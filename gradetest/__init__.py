# -*- coding: utf-8 -*-
# gradetest adapted from SDAPS code by Terry Kovacs terrence.kovacs@dartmouth.edu
# SDAPS - Scripts for data acquisition with paper based surveys
# Copyright(C) 2007-2008, Christoph Simon <post@christoph-simon.eu>
# Copyright(C) 2007-2008, Benjamin Berg <benjamin@sipsolutions.net>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#from gtk._gtk import TREE_SORTABLE_UNSORTED_SORT_COLUMN_ID

u"""
This modules contains a GTK+ user interface for scoreing "bubble test" forms.
"""

from gi import require_version
require_version('Gtk', '3.0')
from gi.repository import GObject
from gi.repository import GLib
from gi.repository import Gtk
from gi.repository import Gdk
import os
import time
import datetime
import sys
import signal
import string
import shutil
import re

from sdaps import model
from sdaps import surface
from sdaps import clifilter
from sdaps import defs
from sdaps import paths
from sdaps import log
from sdaps.add import add_image
from sdaps import  image
from sdaps import setuptex
from sdaps import matrix
from sdaps.utils.ugettext import ugettext, ungettext
_ = ugettext

from exifread import process_file

from sdaps.gui.sheet_widget import SheetWidget
import sdaps.recognize.buddies
import sdaps.gui.buddies
import sdaps.gui.widget_buddies
import savepdfcsv
import shared
import ConfigParser

zoom_steps = [0.15, 0.18, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0,
              1.25, 1.5, 2.0, 2.5, 3.0]

#gain extra page space - move the default sdaps corner marks closer to the edge of the page
defs.corner_mark_left = 6.0 # mm
defs.corner_mark_right = 6.0 # mm
defs.corner_mark_top = 6.0 # mm
defs.corner_mark_bottom = 6.0 # mm
# loosen up the corner box coverage requirements
# The coverage above which a cornerbox is considered to be a logical 1
defs.cornerbox_on_coverage = 0.4

class GradeSheetWidget(SheetWidget):
    def do_button_press_event(self, event):
        # Pass everything except normal clicks down
        if event.button != 1 and event.button != 2 and event.button != 3:
            return False

        if event.button == 2:
            self._drag_start_x = event.x
            self._drag_start_y = event.y
            cursor = Gdk.Cursor(Gdk.CursorType.HAND2)
            self.get_window().set_cursor(cursor)
            return True

        mm_x, mm_y = self._widget_to_mm_matrix.transform_point(event.x, event.y)

        if event.button == 3:
            # Give the corresponding widget the focus.
            box = self.provider.survey.questionnaire.gui.find_box(self.provider.image.page_number, mm_x, mm_y)
            if hasattr(box, "widget"):
                box.widget.focus()

            return True

        # button 1
        self.grab_focus()

        # Look for edges to drag first(on a 4x4px target)
        tollerance_x, tollerance_y = self._widget_to_mm_matrix.transform_distance(4.0, 4.0)
        result = self.provider.survey.questionnaire.gui.find_edge(self.provider.image.page_number, mm_x, mm_y,
                                                                  tollerance_x, tollerance_y)
        if result:
            self._edge_drag_active = True
            self._edge_drag_obj = result[0]
            self._edge_drag_data = result[1]
            return True

        box = self.provider.survey.questionnaire.gui.find_box(self.provider.image.page_number, mm_x, mm_y)
        if box is not None:
            box.data.state = not box.data.state
            shared.grades.score_by_page(self.provider.index)  # rescore the page
            return True    
    
def gui(survey, grades):
    Filter = clifilter.clifilter(survey, ' ')
    provider = Provider(survey, Filter)

    shared.grades = grades
    shared.provider = provider

    try:
        # Exit the mainloop if Ctrl+C is pressed in the terminal.
        GLib.unix_signal_add_full(GLib.PRIORITY_HIGH, signal.SIGINT, lambda *args : Gtk.main_quit(), None)
    except AttributeError:
        # Whatever, it is only to enable Ctrl+C anyways
        pass

    MainWindow(provider).run()

class Grades(object):
    def __init__(self, test):
        self.test = test
        # results by byname
        self.test.byName = {}
        #array of names in the order they were scanned
        self.test.names = []
    
        
    def getname(self, boxes):
        firsti = ' '
        lastlist = list('     ') # will be replaced by the first five of the last name
        alphabet = list(string.ascii_lowercase)
        if len(boxes) == 0:
            return ["bscan", "-"]
        for char in alphabet:
            nl = 0
            for nl in [0,1,2,3,4]:
    #            c = boxes.pop(0)
                if boxes.pop(0) and lastlist[nl] == ' ':
                    lastlist[nl] = char
            if boxes.pop(0) and firsti == ' ': # first initial
                firsti = char
        last = ''
        for c in lastlist:
            last += c
        return [last, firsti]
    
    def gettest_name_answers(self):
        """
        get the name and answers from the sheet boxes
        boxIdNameIndex = 1    #section one of the tex output is the name
        boxIdQuestionIndex = 2 #section two (or greater) of the tex output are questions
        """    
        boxIdNameIndex = 1  # section one of the tex output
        boxIdQuestionIndex = 2  # section two of the tex output
    
        nameBoxes = []
        questionBoxes = []
        quality = 1.0
        test = self.test
        for qobject in test.questionnaire.qobjects:
            if isinstance(qobject, model.questionnaire.Question):
                # Only print data if an image for the page has been loaded
                if test.sheet.get_page_image(qobject.page_number) is None:
                    continue
                for box in qobject.boxes:
                    quality = min(quality, box.data.quality)
                    if box.id[0] == boxIdNameIndex:
                        nameBoxes.append(box.data.state)
                    elif box.id[0] >= boxIdQuestionIndex:
                        questionBoxes.append(box.data.state)
                    else:
                        print 'unknown section: ' + '_'.join([str(num) for num in box.id])
        [last, fi] = self.getname(nameBoxes)
        fullName = last + ' ' + fi
        # dis-ambiguate name
        dup = 2
        while fullName in test.byName:
            fullName = last + ' ' + fi + '(' + str(dup) + ')'
            dup += 1
        test.questionnaire.fullName = fullName
        test.questionnaire.quality = quality
        test.questionnaire.score = 0
        test.questionnaire.answers = []
        test.names.append(fullName)
        test.byName[fullName] = {'answers':[], 'marked_answer':[], 'quality':quality, 'score':0, 'page_index':test.index }
    #        test.ByName[fullName] = quality
        while len(questionBoxes) >= 5:
            answer = ''
            for ans in list('abcde'):
                if questionBoxes.pop(0):
                    answer += ans  # there may be more than one correct answer to a question
            if answer == '':  # no answer was given
                answer = '-'
            test.byName[fullName]['answers'].append(answer)
            test.questionnaire.answers.append(answer)
        return fullName
        
    def score(self):
        """
        score all sheets in the test
        """
        test = self.test
        try:
            answers = test.byName['answe r']['answers'] # find the answer key
        except:
            for name in test.byName:
                test.byName[name]['score'] = 0
            return False

        for name in test.byName:
            ncorrect = 0
            index = 0
            if len(test.byName[name]['answers']) > 0:
                test.byName[name]['marked_answer'] = []
                for answer in answers:
                    if answer != "-" and test.byName[name]['answers'][index] in answer:
                        ncorrect += 1
                        test.byName[name]['marked_answer'].append(str.upper(test.byName[name]['answers'][index]))
                    else:
                        test.byName[name]['marked_answer'].append(str.lower(test.byName[name]['answers'][index]))
                    index += 1
            test.byName[name]['score'] = ncorrect
        return True
    
    def score_by_name(self, name):
        """
        name = name of sheet to score
        score named sheet
        """
        test = self.test
        shared.project_saved = False        
        ncorrect = 0
        index = 0
        try:
            answers = test.byName['answe r']['answers'] # find the answer key
        except:
            return ncorrect # no answer key
        test.byName[name]['marked_answer'] = []
        for answer in answers:
            if answer != "-" and test.byName[name]['answers'][index] in answer:
                ncorrect += 1
                test.byName[name]['marked_answer'].append(str.upper(test.byName[name]['answers'][index]))
            else:
                test.byName[name]['marked_answer'].append(str.lower(test.byName[name]['answers'][index]))
            index += 1
        test.byName[name]['score'] = ncorrect

        return ncorrect
    
    def score_by_page(self, page_number):
        """
        score one sheet in the test
        """
        test = self.test
        for name in test.byName:
            if test.byName[name]['page_index'] == page_number + 1:
                test.index = page_number + 1
                del test.byName[name]
                name = self.gettest_name_answers()
                score = self.score_by_name(name)
                name_store = shared.provider.main_window.name_store
                name_iter = shared.provider.main_window.name_store_byPage[page_number]
                name_store[name_iter][0] = name
                name_store[name_iter][1] = score
                break

        if name == 'answe r':  #  answer key changed - rescore all test sheets
            self.score()
            for name in test.byName:
                page_number = test.byName[name]['page_index'] - 1
                name_store = shared.provider.main_window.name_store
                name_iter = shared.provider.main_window.name_store_byPage[page_number]
#                name_store[name_iter][0] = name
                name_store[name_iter][1] = test.byName[name]['score']                
        return

class Provider(object):

    def __init__(self, survey, filter, by_quality=False):
        self._by_quality = by_quality

        self.survey = survey
        self.images = list()
        self.qualities = list()
        self.survey.iterate(self, filter)
        self.qualities.sort(reverse=False)
        self.index = 0

        # There may be no images. This error is
        # caught and printed in the "gui" function.
        if not self.images:
            return

        self.image.surface.load_rgb()
        self.survey.goto_sheet(self.image.sheet)
        #self._surface = None

    def __call__(self):
        # Add all images that are "valid" ie. everything except back side of
        # a simplex printout
        new_images = [img for img in sorted(self.survey.sheet.images, key=lambda Image: Image.page_number) if not img.ignored]

        self.images.extend(new_images)
        # Insert each image of the sheet into the qualities array
        for i in xrange(len(new_images)):
            self.qualities.append((self.survey.sheet.quality, len(self.qualities)))

    def next(self, cycle=True):
        if self.index >= len(self.images) - 1:
            if cycle:
                new_index = 0
            else:
                return False
        else:
            new_index = self.index + 1
        self.image.surface.clean()

        self.index = new_index

        self.image.surface.load_rgb()
        self.survey.goto_sheet(self.image.sheet)
        name_iter = self.main_window.name_store_byPage[self.index]
        self.main_window.nameSelect.select_iter(name_iter)

        return True

    def previous(self, cycle=True):
        if self.index <= 0:
            if cycle:
                new_index = len(self.images) - 1
            else:
                return False
        else:
            new_index = self.index - 1

        self.image.surface.clean()

        self.index = new_index

        self.image.surface.load_rgb()
        self.survey.goto_sheet(self.image.sheet)
        name_iter = self.main_window.name_store_byPage[self.index]
        self.main_window.nameSelect.select_iter(name_iter)

        return True

    def goto(self, index):
        if index >= 0 and index < len(self.images):
            self.image.surface.clean()
            self.index = index
            self.image.surface.load_rgb()
            self.survey.goto_sheet(self.image.sheet)
            name_iter = self.main_window.name_store_byPage[index]
            self.main_window.nameSelect.select_iter(name_iter)

    def set_sort_by_quality(self, value):
        self.image.surface.clean()
        self._by_quality = value
        self.image.surface.load_rgb()
        self.survey.goto_sheet(self.image.sheet)

    def get_image(self):
        if self._by_quality:
            return self.images[self.qualities[self.index][1]]
        else:
            return self.images[self.index]

    image = property(get_image)

def cygwin_connect_signals(builder, sigdict):
    for objkey in sigdict:        
        event, app = sigdict[objkey]
        sigobject = builder.get_object(objkey)
        sigobject.connect(event, app)
    return
    
class GradeSetupWindow(object):
    """
    Called w/o arguments - test tif file and form template collected from fields in window
    """

    def __init__(self):
        self.scan_file = None
        self.grades_dir = None  # top level directory eg $HOME/Grades
        self.template_dir = None  # the default place for templates to be found self.grades_dir/templates
        self.project_dir = None  # default place for the test surveys - self.grades_dir/projects
        self.test_dir = None  # the actual test directory - self.projects_dir/???
        self.test_dir_userset = False
        self.provider = None
        self.about_dialog = None
        self._builder = Gtk.Builder()
        setup_window_signals = {"main_window": ("delete-event", self.quit_application),
                                "begin": ("clicked", self.begin_cb),
                                "exit": ("clicked", self.quit_application),
                                "imagemenuitem10" : ("activate", self.show_about_dialog),
                                "template_chooser_button": ("clicked",  self.template_dir_cb),
                                "project_chooser_button": ("clicked", self.project_dir_cb),
                                "scanfile_chooser_button": ("clicked", self.scan_file_cb)}

        paths.init(False, __path__[0])
        if paths.local_run:
            self._builder.add_from_file(
                os.path.join(os.path.dirname(__file__), 'grade_setup_window.ui'))
        else:
            self._builder.add_from_file(
                os.path.join(
                    paths.prefix,
                    'share', 'gradetest', 'ui', 'grade_setup_window.ui'))
        self._window = self._builder.get_object("main_window")
        cygwin_connect_signals(self._builder, setup_window_signals)
#        self._builder.connect_signals(self)
        self._window.show()
        self.get_defaults()
        self.template_choose_button = self._builder.get_object("template_chooser_button")
        
        self.scanfile_choose_button = self._builder.get_object("scanfile_chooser_button")

        self.project_choose_button = self._builder.get_object("project_chooser_button")

        Gtk.main()
        
    def get_defaults(self):
        """
        read from config file or create a config file
        """
        config = ConfigParser.RawConfigParser()
        try:
            configfilepath = os.path.expanduser('~/') + shared.config_file
            config.read(configfilepath)
            file_configs = config.items('grades')
            for entry in file_configs:
                (key, value) = entry
                if value != 'None' :
                    shared.configs[key] = value
            self.grades_dir = shared.configs['grades directory']
            self.template_dir = shared.configs['template directory']
            self.project_dir = shared.configs['project directory']
                        
        except:
            self.grades_dir = shared.configs['grades directory']
            self.template_dir = shared.configs['template directory']
            self.project_dir = shared.configs['project directory']
            try:
                configfile = open(configfilepath, 'w')
                config.add_section('grades')
                shared.configs['grades directory'] = self.grades_dir
                shared.configs['template directory'] = self.template_dir
                shared.configs['project directory'] = self.project_dir
                for key in shared.configs:
                    config.set('grades', key, shared.configs[key])
                config.write(configfile)
            except:
                shared.popup(self._window, "Error: cannot create " + configfilepath)
            
        try:
            os.stat(self.template_dir)
        except:
            try:
                os.makedirs(self.template_dir, mode=0750)
            except:
                shared.popup(self._window, "Error: cannot create " + self.template_dir)               
        try:
            os.stat(self.project_dir)
        except:
            try:
                os.makedirs(self.project_dir, mode=0750)
            except:
                shared.popup(self._window, "Error: cannot create " + self.project_dir)
        return
   
    def show_about_dialog(self, *args):
        if not self.about_dialog:
            self.about_dialog = Gtk.AboutDialog()
            self.about_dialog.set_logo(None)
            self.about_dialog.set_program_name("GradeTest")
            self.about_dialog.set_version("1.0 - Beta")
            self.about_dialog.set_authors(
                [u"Terrence Kovacs <Terrence.Kovacs@dartmouth.edu>"])
            self.about_dialog.set_copyright(_(u"Copyright © 2016"))
            self.about_dialog.set_license_type(Gtk.License.GPL_3_0)
            self.about_dialog.set_comments(u"Automated grading of multiple choice tests\n based on code from http://sdaps.org")
            #self.about_dialog.set_website(_(u"http://sdaps.org"))
            #self.about_dialog.set_translator_credits(_("translator-credits"))
            self.about_dialog.set_default_response(Gtk.ResponseType.CANCEL)

        self.about_dialog.set_transient_for(self._window)
        self.about_dialog.run()
        self.about_dialog.hide()

        return True
        
    def template_dir_cb(self, widget):
        dialog = Gtk.FileChooserDialog("Choose a test template", self._window,
            Gtk.FileChooserAction.SELECT_FOLDER,
            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
             "Select", Gtk.ResponseType.OK))
        dialog.set_default_size(400, 400)
        dialog.set_current_folder(self.template_dir)
        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            self.template_dir = dialog.get_filename()
            shared.test_template = os.path.split(self.template_dir)[1]
            widget.set_label(shared.test_template)
        elif response == Gtk.ResponseType.CANCEL:
            dialog.destroy()
            return       
        dialog.destroy()
        
        if shared.test_scanfile != None :  # we should have everything we need now
            self.init_project()
        return
    
    def project_dir_cb(self, widget):
        dialog = Gtk.FileChooserDialog("Choose a project", self._window,
            Gtk.FileChooserAction.SELECT_FOLDER,
            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
             "Select", Gtk.ResponseType.OK))
        dialog.set_default_size(800, 400)
        dialog.set_current_folder(self.project_dir)

        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            self.test_dir = dialog.get_filename()
            self.test_dir_userset = True
            widget.set_label(os.path.split(self.test_dir)[1])            
        elif response == Gtk.ResponseType.CANCEL:
            dialog.destroy()
            return
        
        dialog.destroy()
        import glob
        # look for template and scanfile if they are not defined
        if shared.test_template == None:
            try:
                shared.test_template = glob.glob(self.test_dir + '/*.tex')[0]
                shared.test_template = os.path.splitext(shared.test_template)[0]
                shared.test_template = os.path.split(shared.test_template)[1]
                self.template_choose_button.set_label(shared.test_template)
            except:
                shared.test_template = None
        if shared.test_scanfile == None:
            try:
                shared.test_scanfile = glob.glob(self.test_dir + '/*.tif')[0]
                shared.test_scanfile = os.path.split(shared.test_scanfile)[1]
                self.scanfile_choose_button.set_label(shared.test_scanfile)
            except:
                shared.test_scanfile = None
        if shared.test_scanfile == None or shared.test_template == None :  # not a previous project
            return
        try:
            os.stat(self.test_dir + '/survey')
        except:
            return  # not an existing project
        os.chdir(self.test_dir)
        test = model.survey.Survey.load(self.test_dir)
        grades = Grades(test)
        shared.grades = grades
        grades.test_dir = self.test_dir #  so we know where a future pdf/csv file might go
        for page in range(1, len(test.sheets)):
            test.index = page
            grades.gettest_name_answers()            
        grades.score()
        self._window.destroy()
        gui(test, grades)

        return
    
    def scan_file_cb(self, widget):
        dialog = Gtk.FileChooserDialog("Choose a test scan file", self._window,
            Gtk.FileChooserAction.OPEN,
            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
             "Select", Gtk.ResponseType.OK))
        dialog.set_default_size(800, 400)
        dialog.set_current_folder(shared.configs['scan directory'])
        scanfile_filter = Gtk.FileFilter()
        scanfile_filter.add_pattern("*.tif")
        dialog.set_filter(scanfile_filter)

        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            self.scan_file = dialog.get_filename()
            ( shared.configs['scan directory'],  shared.test_scanfile ) = os.path.split(self.scan_file)
            widget.set_label(shared.test_scanfile)                        
        elif response == Gtk.ResponseType.CANCEL:
            dialog.destroy()
            return

        dialog.destroy()
        
        if shared.test_template != None : # we should have everyting we need now
            self.init_project()
        return

    def init_project(self):
        """
        create project directory ddmmyy/scanfile if none exists
        """
        if self.test_dir_userset :
            return
                
        if self.test_dir == None :
            scan_part = os.path.splitext(self.scan_file)[0]
            scan_part = os.path.basename(scan_part)
            today = datetime.datetime.today()
            date_string = today.strftime('%d%m%y')
            self.test_dir = self.project_dir + '/' + date_string + '/' + scan_part
        
        try:
            os.stat(self.test_dir)
            overwrite = shared.popup_yesno(self._window,self.test_dir + ' exists, overwrite?')
            if overwrite == Gtk.ResponseType.YES:
                os.chdir(self.project_dir + '/' + date_string)
                shutil.rmtree(scan_part)
                os.makedirs(self.test_dir, 0750)
            else:
                self.quit_application() # to do - something better here
                self.test_dir = None
                self.project_choose_button.set_label("")
                return
        except:
            try:
                os.makedirs(self.test_dir, 0750)
            except:
                shared.popup(self._window,'Error - unable to create project directory ' + self.test_dir)
                return
        self.project_choose_button.set_label(date_string + '/' + scan_part)
        return
        

    def begin_cb(self, widget):
        try:
            # existing project
            os.stat(self.test_dir + '/questionnaire.sdaps')
            test = model.survey.Survey.load(self.test_dir)
            grades = Grades(test)
            shared.grades = grades
            for page in range(1, len(test.sheets)):
                test.index = page
                grades.gettest_name_answers(test)
        except:
            # new project
            if self.test_dir == None or self.scan_file == None :
                shared.popup(self._window,'Need template and scanfile or existing project')
                return
            for name in os.listdir(self.template_dir):
                shutil.copy(self.template_dir + '/' + name, self.test_dir)
            shutil.copy(self.scan_file, self.test_dir)
            scan_file_name = os.path.split(self.scan_file)[1]
            self.scan_file = os.path.join(self.test_dir, scan_file_name)
            # create new survey to store the test data
            test = model.survey.Survey.new(self.test_dir)
            test.add_questionnaire(model.questionnaire.Questionnaire())
            # add checkbox info from questionaire.sdaps file
            setuptex.sdapsfileparser.parse(test)
            sheet = model.sheet.Sheet()
            test.add_sheet(sheet)
            test.sheet.images = []
            # add sheet images from the scan file
            add_image(test, self.scan_file, test.defs.duplex, copy=False)
            # a place to store the grades
            grades = Grades(test)
            shared.grades = grades
            num_pages = image.get_tiff_page_count(self.scan_file)
            # find checkmarks on each page and collect the names and answers
            for page in range(num_pages):
                test.index = page + 1
                test.questionnaire.recognize.recognize()
                grades.gettest_name_answers()
        grades.test_dir = self.test_dir
        # score the test
        grades.score()
        self._window.destroy()
        # open the correction/verification window
        gui(test, grades)
        return

    
    def quit_application(self, *args):
        Gtk.main_quit()
        self._window.destroy()
        exit
        
        
    def run(self):
        self._window.show()
        Gtk.main()


class MainWindow(object):
    def __init__(self, provider):
#        self.scanFile = None
#        self.templateFile = None
        main_window_signals = {"OK Button" : ("clicked", self.save_cancel_ok_cb),
                               "Cancel Button" : ("clicked", self.save_cancel_ok_cb),
                               "save_filename" : ("clicked", self.save_filename_cb),
                               "action-next-page" : ("activate", self.go_to_next_page),
                               "action-prev-page" : ("activate", self.go_to_previous_page),
                               "zoom-in" : ("activate", self.zoom_in),
                               "zoom-out" : ("activate", self.zoom_out),
                               "main_window" : ("delete-event", self.quit_application),
                               "imagemenuitem3" : ("activate", self.save_project),
                               "imagemenuitem1" : ("activate", self.save_csv_cb),
                               "imagemenuitem2" : ("activate", self.save_pdf_cb),
                               "imagemenuitem5" : ("activate", self.quit_application),
                               "imagemenuitem10" : ("activate", self.show_about_dialog),
                               "save" : ("clicked", self.save_project),
                               "verifyonview" : ("toggled", self.verifyonview_toggled_cb),
                               "fullscreen_toolbutton" : ("clicked", self.toggle_fullscreen),
                               "page_spin" : ("value-changed", self.page_spin_value_changed_cb)}
#                            "turned_toggle" : ("toggled", self.turned_toggle_toggled_cb),
#                            "rescore" : ("clicked", self.rescore_cb)}

        self.about_dialog = None
        self.close_dialog = None
#        self.ask_open_dialog = None
        self.saved = False  # have changes been saved?
        self.verifyonview = False
        self.provider = provider
        provider.main_window = self  # link so others can access my attributes/methods
        self.name_store_byPage = []   # list of storeitems by page index
        self._load_image = 0
        self._builder = Gtk.Builder()
        if paths.local_run:
            self._builder.add_from_file(
                os.path.join(os.path.dirname(__file__), 'grade_window.ui'))
        else:
            self._builder.add_from_file(
                os.path.join(
                    paths.prefix,
                    'share', 'gradetest', 'ui', 'grade_window.ui'))

        self._window = self._builder.get_object("main_window")
        shared.mainwindow = self._window
        cygwin_connect_signals(self._builder, main_window_signals)
#       self._builder.connect_signals(self)
        self._window.resize(int(shared.configs['mainwindow width']), int(shared.configs['mainwindow height']))
        if shared.configs['mainwindow x'] != None and shared.configs['mainwindow y'] != None:
            self._window.move(int(shared.configs['mainwindow x']), int(shared.configs['mainwindow y']))

        template_text_widget = self._builder.get_object("template")
        template_text_widget.set_text(shared.test_template)
        
        # try to find the scan date
        tiff_file = open(shared.grades.test_dir + '/' + shared.test_scanfile, 'rb')
        tags = process_file(tiff_file)
        shared.scan_date = 'Unknown'
        for key in tags:
            if re.search('[Dd]ate[Tt]ime', key):
                shared.scan_date = tags[key]
                break
            
        scanfile_text_widget = self._builder.get_object("scanfile")
        scanfile_text_widget.set_text(shared.test_scanfile + " Date Scanned: " + str(shared.scan_date) )
                
        scrolled_window = self._builder.get_object("sheet_scrolled_window")
        self.sheet = GradeSheetWidget(self.provider)
        self.sheet.show()
        scrolled_window.add(self.sheet)

        self.data_viewport = self._builder.get_object("data_view")
        widgets = provider.survey.questionnaire.widget.create_widget()
        widgets.show_all()

        self.data_viewport.add(widgets)

        provider.survey.questionnaire.widget.connect_ensure_visible(self.data_view_ensure_visible)

        self.sheet.connect('key-press-event', self.sheet_view_key_press)

        self.name_view = self._builder.get_object("name_view")
        self.name_store = Gtk.ListStore(str,int,str,int,bool)
        for name in provider.survey.names:
            attributes = provider.survey.byName[name]
            provider.survey.index = attributes['page_index']
            storeitem = self.name_store.append([name, attributes['score'], '{:.2f}'.format(attributes['quality']), 
                        attributes['page_index'], provider.survey.sheet.verified])
            self.name_store_byPage.insert(attributes['page_index'], storeitem)
        provider.survey.index = 1  # reset to first page
        self.nameView = Gtk.TreeView(self.name_store)
        self.name_view.add(self.nameView)
        
        self.nameSelect = self.nameView.get_selection()
        self.nameSelect.connect("changed", self.nameSelectionChanged)
        
        rendererText = Gtk.CellRendererText()
        column = Gtk.TreeViewColumn("Name", rendererText, text=0)
        column.set_sort_column_id(0)
        self.name_store.set_sort_func(0, self.sort_names, None)
        column.set_property('sizing',Gtk.TreeViewColumnSizing.AUTOSIZE)   
        self.nameView.append_column(column)
        
        rendererGrade = Gtk.CellRendererText()
        rendererGrade.set_property('xalign',1.0)
        rendererGrade.props.width_chars = 5
        column = Gtk.TreeViewColumn("Grade", rendererGrade, text=1)
        column.set_sort_column_id(1)
        self.name_store.set_sort_func(1, self.sort_names, None)
        column.set_property('sizing',Gtk.TreeViewColumnSizing.AUTOSIZE)
        self.nameView.append_column(column)

        rendererQuality = Gtk.CellRendererText()
        rendererQuality.set_property('xalign', 0.5)
        column = Gtk.TreeViewColumn("Quality", rendererQuality, text=2)
#        column.set_property('fixed-width', 60)
        column.set_sort_column_id(2)
        self.name_store.set_sort_func(2, self.sort_names, None)
        column.set_property('sizing',Gtk.TreeViewColumnSizing.AUTOSIZE)
        self.nameView.append_column(column)
        
        rendererPage = Gtk.CellRendererText()
        rendererPage.set_property('xalign', 0.5)
        column = Gtk.TreeViewColumn("Page", rendererPage, text=3)
#        column.set_property('fixed-width', 50)
        column.set_sort_column_id(3)
        self.name_store.set_sort_func(3, self.sort_names, None)
        column.set_property('sizing',Gtk.TreeViewColumnSizing.AUTOSIZE)
        self.nameView.append_column(column)
        
        rendererVerified = Gtk.CellRendererToggle()
        rendererVerified.connect('toggled', self.verified_cb)
        rendererVerified.set_property('xalign', 0)
        rendererVerified.set_property('xpad', 10)
        column = Gtk.TreeViewColumn("Verified", rendererVerified, active = 4)
        column.set_sort_column_id(4)
        self.name_store.set_sort_func(4, self.sort_names, None)
        column.set_property('sizing',Gtk.TreeViewColumnSizing.AUTOSIZE)    
        self.nameView.append_column(column)
#        self.nameView.AutoSizeColumns()
        self.nameView.show_all()
        
        # used for saving bot PDF and CSV files
        self.save_file_dialog = self._builder.get_object("save_file_dialog")
        self.save_filename = self._builder.get_object('save_filename')
        self.supporting_info = self._builder.get_object('supporting_info')
        self.save_file_dialog.set_transient_for(self._window)
        if shared.configs["pdf directory"] == None:
            shared.configs["pdf directory"] = shared.grades.test_dir    
        savepdfcsv.pdf_filename = shared.configs["pdf directory"] + '/grades.pdf'
        if shared.configs["csv directory"] == None:
            shared.configs["csv directory"] = shared.grades.test_dir    
        savepdfcsv.csv_filename = shared.configs["csv directory"] + '/grades.csv'
    
        self.sheet.props.zoom = float(shared.configs['zoom'])
        self.save_project()
        shared.project_saved = True
        self.update_ui()
        return

#============================================================================
# custom sort function keeps the answer key at the top
#============================================================================
    def sort_names(self, tree, row1, row2, user_data):
        sort_column, _ = tree.get_sort_column_id()
        column = self.nameView.get_column(sort_column)
        sort_order = column.get_sort_order()
        name1 = tree.get_value(row1, 0)
        name2 = tree.get_value(row2, 0)
        if sort_order == Gtk.SortType.ASCENDING:
            if name1 == 'answe r': return -1
            if name2 == 'answe r': return 1
        else:
            if name1 == 'answe r': return 1
            if name2 == 'answe r': return -1
        value1 = tree.get_value(row1, sort_column)
        value2 = tree.get_value(row2, sort_column)
        if value1 < value2:
            return -1
        elif value1 == value2:
            return 0
        else:
            return 1
        
    def save_pdf_cb(self,widget):
        self.save_file_dialog.set_title('Grades to PDF')
        filename = os.path.split(savepdfcsv.pdf_filename)[1]
        self.save_filename.set_label(savepdfcsv.pdf_filename)
#        sbuffer = self.supporting_info.get_buffer()
#        sbuffer.set_text(savepdfcsv.supporting_info, len(savepdfcsv.supporting_info))
        savepdfcsv.do_pdf = True
        self.save_file_dialog.run()
        
    def save_csv_cb(self,widget):
        self.save_file_dialog.set_title('Grades to CSV')
        filename = os.path.split(savepdfcsv.csv_filename)[1]
        self.save_filename.set_label(savepdfcsv.csv_filename)
        savepdfcsv.do_pdf = False
        self.save_file_dialog.run()
        
    def save_filename_cb(self, widget):
        dialog = Gtk.FileChooserDialog("Choose a file", self._window,
            Gtk.FileChooserAction.SAVE,
            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
             "Select", Gtk.ResponseType.OK))
        if savepdfcsv.do_pdf:
            dwhere, filename = os.path.split(savepdfcsv.pdf_filename)
        else:
            dwhere, filename = os.path.split(savepdfcsv.csv_filename)
        dialog.set_current_name(filename)
        dialog.set_current_folder(dwhere)
        dialog.set_default_size(400, 400)

        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            filepath = dialog.get_filename()
            dwhere, filename = os.path.split(filepath)
            self.save_filename.set_label(filepath)
            if savepdfcsv.do_pdf: 
                savepdfcsv.pdf_filename = filepath
                if dwhere != shared.configs['pdf directory']:
                    shared.project_saved = False
                shared.configs['pdf directory'] = dwhere
            else:            
                savepdfcsv.csv_filename = filepath
                if dwhere != shared.configs['csv directory']:
                    shared.project_saved = False
                shared.configs['csv directory'] = dwhere
#        elif response == Gtk.ResponseType.CANCEL:
#            print("Cancel clicked")
        
        dialog.destroy()
        return
        
    def save_cancel_ok_cb(self, widget):
        self.save_file_dialog.hide()
        if widget.get_name() != 'save_ok':
            return
        
        savepdfcsv.supporting_info  = self.supporting_info.props.buffer.props.text
        savepdfcsv.name_store = self.name_store
        savepdfcsv.save()
        return
        
    def verified_cb(self, widget, path):
        """
        column 4 -> verified true/false
        column 3 -> sheet index
        """
        self.name_store[path][4] = not self.name_store[path][4]
        self.provider.survey.index = self.name_store[path][3]
        self.provider.survey.sheet.verified = self.name_store[path][4]
        shared.project_saved = False
        return
    
    def nameSelectionChanged(self, selection):
        model, index = selection.get_selected()
        if index != None:
            name = model[index][0]
            # image pages start at 0 and test pages start at 1
            newpage = self.provider.survey.byName[name]['page_index'] - 1

            self.go_to_page(newpage)
            
    def zoom_in(self, *args):
        cur_zoom = self.sheet.props.zoom
        try:
            i = zoom_steps.index(cur_zoom)
            i += 1
            if i < len(zoom_steps):
                self.sheet.props.zoom = zoom_steps[i]
        except:
            self.sheet.props.zoom = 1.0

    def zoom_out(self, *args):
        cur_zoom = self.sheet.props.zoom
        try:
            i = zoom_steps.index(cur_zoom)
            i -= 1
            if i >= 0:
                self.sheet.props.zoom = zoom_steps[i]
        except:
            self.sheet.props.zoom = 1.0

    def null_event_handler(self, *args):
        return True

    def show_about_dialog(self, *args):
        if not self.about_dialog:
            self.about_dialog = Gtk.AboutDialog()
            self.about_dialog.set_program_name("GradeTest")
            #self.about_dialog.set_version("")
            self.about_dialog.set_authors(
                [u"Terrence Kovacs <Terrence.Kovacs@dartmouth.edu>"])
            self.about_dialog.set_copyright(_(u"Copyright © 2016"))
            self.about_dialog.set_license_type(Gtk.License.GPL_3_0)
            self.about_dialog.set_comments(u"Automated grading of multiple choice tests\n based on code from http://sdaps.org")

            #self.about_dialog.set_website(_(u"http://sdaps.org"))
            #self.about_dialog.set_translator_credits(_("translator-credits"))
            self.about_dialog.set_default_response(Gtk.ResponseType.CANCEL)

        self.about_dialog.set_transient_for(self._window)
        self.about_dialog.run()
        self.about_dialog.hide()

        return True

    def update_page_status(self):
#        combo = self._builder.get_object("page_number_combo")
        turned_toggle = self._builder.get_object("turned_toggle")

        # Update the combobox
        if self.provider.image.survey_id == self.provider.survey.survey_id:
            page_number = self.provider.image.page_number
        else:
            page_number = -1

        # Update the toggle
#        turned_toggle.set_active(self.provider.image.rotated or False)

    def update_ui(self):

        position_label = self._builder.get_object("position_label")
        quality_label = self._builder.get_object("quality_label")
        page_spin = self._builder.get_object("page_spin")
        position_label.set_text(_(u" of %i") % len(self.provider.images))
        quality_label.set_text(_(u"Recognition Quality: %.2f") % self.provider.image.sheet.quality)
        #position_label.props.sensitive = True
        page_spin.set_range(1, len(self.provider.images))
        page_spin.set_value(self.provider.index + 1)

        self.update_page_status()
        self.sheet.update_state()

        self.provider.survey.questionnaire.widget.sync_state()

    def go_to_previous_page(self, *args):
        if not self.provider.previous(cycle=False):
            dialog = Gtk.MessageDialog(
                flags=Gtk.DialogFlags.MODAL | Gtk.DialogFlags.DESTROY_WITH_PARENT,
                type=Gtk.MessageType.INFO,
                buttons=Gtk.ButtonsType.CANCEL,
                message_format=_("You have reached the first page of the survey. Would you like to go to the last page?"))
            dialog.set_transient_for(self._window)
            dialog.add_button(_("Go to last page"), Gtk.ResponseType.OK)
            if dialog.run() == Gtk.ResponseType.OK:
                self.provider.previous(cycle=True)
            dialog.destroy()
        self.update_ui()
        return True

    def go_to_page(self, page):
        if page == self.provider.index:
            return True
        if self.verifyonview and not self.provider.image.sheet.verified:
                self.provider.image.sheet.verified = True
                path = self.name_store_byPage[self.provider.index]
                self.name_store[path][4] = True  # sets checkmark in the name view
                # Mark the sheet as valid
                self.provider.image.sheet.valid = True
                shared.project_saved = False
                
        self.provider.goto(int(page))

        self.update_ui()
        return True

    def go_to_next_page(self, *args):
        if not self.provider.next(cycle=False):
            dialog = Gtk.MessageDialog(
                flags=Gtk.DialogFlags.MODAL | Gtk.DialogFlags.DESTROY_WITH_PARENT,
                type=Gtk.MessageType.INFO,
                buttons=Gtk.ButtonsType.CANCEL,
                message_format=_("You have reached the last page of the survey. Would you like to go to the first page?"))
            dialog.set_transient_for(self._window)
            dialog.add_button(_("Go to first page"), Gtk.ResponseType.OK)
            if dialog.run() == Gtk.ResponseType.OK:
                self.provider.next(cycle=True)
            dialog.destroy()

        self.update_ui()
        return True

    def page_spin_value_changed_cb(self, *args):
        page_spin = self._builder.get_object("page_spin")
        page = page_spin.get_value() - 1
        self.go_to_page(page)

    def turned_toggle_toggled_cb(self, *args):
        toggle = self._builder.get_object("turned_toggle")
        rotated = toggle.get_active()
        if self.provider.image.rotated != rotated:
            self.provider.image.rotated = rotated
            self.provider.image.surface.load_rgb()
#            self.provider.survey.questionnaire.recognize.recognize()
            self.update_ui()
        return False

    def verifyonview_toggled_cb(self, *args):
        toggle = self._builder.get_object("verifyonview")
        self.verifyonview = toggle.get_active()
        return False
    
    def toggle_fullscreen(self, *args):
        flags = self._window.get_window().get_state()
        if flags & Gdk.WindowState.FULLSCREEN:
            self._window.unfullscreen()
        else:
            self._window.fullscreen()
        return True

    def save_project(self, *args):
        shared.project_saved = True
        self.provider.survey.save()
        # save the config file
        shared.configs['zoom'] = self.sheet.props.zoom
        """
        test_dir is default and unique to a project
        remember only if they have been explicitly set - ie "Desktop" or "Documents"
        """
        if shared.configs['csv directory'] == shared.grades.test_dir:
            shared.configs['csv directory'] = None
        if shared.configs['pdf directory'] == shared.grades.test_dir:
            shared.configs['pdf directory'] = None
        shared.configs['mainwindow width'], shared.configs['mainwindow height'] = self._window.get_size()
        shared.configs['mainwindow x'], shared.configs['mainwindow y'] = self._window.get_position()
        config = ConfigParser.RawConfigParser()
        configfilepath = os.path.expanduser('~/') + shared.config_file
        try:
            configfile = open(configfilepath, 'w')
            config.add_section('grades')
            for key in shared.configs:
                config.set('grades', key, shared.configs[key])
            config.write(configfile)
        except:
            shared.popup(self._window, "Error: cannot create " + configfilepath)
        return True

    def data_view_ensure_visible(self, widget):
        allocation = widget.get_allocation()

        vadj = self.data_viewport.props.vadjustment
        lower = vadj.props.lower
        upper = vadj.props.upper
        value = vadj.props.value
        page_size = vadj.props.page_size

        value = max(value, allocation.y + allocation.height - page_size)
        value = min(value, allocation.y)
        value = max(value, lower)
        value = min(value, upper)

        vadj.props.value = value

    def sheet_view_key_press(self, window, event):
        # Go to the next when Enter or Tab is pressed
        enter_keyvals = [Gdk.keyval_from_name(k) for k in ["Return", "KP_Enter", "ISO_Enter"]]
        if event.keyval in enter_keyvals:
            # If "Return" is pressed, then the examiner figured that the data
            # is good.

            # Mark as verified
            if self.verifyonview:
                self.provider.image.sheet.verified = True
                path = self.name_store_byPage[self.provider.index]
                self.name_store[path][4] = True  # sets checkmark in the name view
                # Mark the sheet as valid
                self.provider.image.sheet.valid = True

            if event.state & Gdk.ModifierType.SHIFT_MASK:
                self.go_to_previous_page()
            else:
                self.go_to_next_page()
            return True
        elif event.keyval == Gdk.KEY_Tab or event.keyval == Gdk.KEY_KP_Tab or event.keyval == Gdk.KEY_ISO_Left_Tab:
            # Allow tabbing out with Ctrl
            if event.state & Gdk.ModifierType.CONTROL_MASK:
                return False

            if event.state & Gdk.ModifierType.SHIFT_MASK:
                self.go_to_previous_page()
            else:
                self.go_to_next_page()
            return True

        return False

    def quit_application(self, *args):
        if shared.project_saved :
            Gtk.main_quit()
            return False

        if not self.close_dialog:
            self.close_dialog = Gtk.MessageDialog(
                parent=self._window,
                flags=Gtk.DialogFlags.MODAL,
                type=Gtk.MessageType.WARNING)
            self.close_dialog.add_buttons(
                _(u"Close without saving"), Gtk.ResponseType.CLOSE,
                Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                Gtk.STOCK_SAVE, Gtk.ResponseType.OK)
            self.close_dialog.set_markup(
                _(u"<b>Save the project before closing?</b>\n\nIf you do not save you may loose data."))
            self.close_dialog.set_default_response(Gtk.ResponseType.CANCEL)

        response = self.close_dialog.run()
        self.close_dialog.hide()

        if response == Gtk.ResponseType.CLOSE:
            Gtk.main_quit()
            return False
        elif response == Gtk.ResponseType.OK:
            self.save_project()
            shared.project_saved = True
            Gtk.main_quit()
            return False
        else:
            return True
        
    #===========================================================================
    # def test_template_file_cb(self, widget):
    #     self.templateFile = widget.get_filename()
    #     return
    # 
    # def test_scan_file_cb(self, widget):
    #     self.scanFile = widget.get_filename()
    #     return
    #===========================================================================
    
    def rescore_cb(self, widget):
        shared.grades.score()
        test = shared.grades.test
        for name in test.byName:
            page_number = test.byName[name]['page_index'] - 1
            name_store = shared.provider.main_window.name_store
            name_iter = shared.provider.main_window.name_store_byPage[page_number]
#                name_store[name_iter][0] = name
            name_store[name_iter][1] = test.byName[name]['score']      
        return
        

    def run(self):
        self._window.show()
#        Gtk.main()

