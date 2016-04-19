'''
Created on Dec 17, 2015
a place for shared information
@author: tkovacs
'''
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
import os
from gi import require_version
require_version('Gtk', '3.0')
from gi.repository import GObject
from gi.repository import GLib
from gi.repository import Gtk
from gi.repository import Gdk

test_template = None
test_scanfile = None
project_saved = False
provider = None
grades = None
config_file = ".Grade"
scan_date = "Unknown"

home = os.path.expanduser('~')
# saved in the config file - these are defaults
configs = {'template directory' : home + "/Grades/templates",
            'scan directory' : home,
            'grades directory' : home + "/Grades",
            'csv directory' : home,
            'pdf directory' : home,
            'project directory' : home + "/Grades/projects",
            'zoom' : 0.5,
            'mainwindow width' : 72 * 11,
            'mainwindow height' : 72 * 11,
            'mainwindow x' : None,
            'mainwindow y' : None }


#configs['grades directory'] = home + "/Grades"
#configs['template directory'] = home + "/Grades/templates"
#configs['project directory'] = home + "/Grades/projects"
#configs['scan directory'] = home

def popup(parent_window, message):
    popup_dialog = Gtk.MessageDialog(
        parent=parent_window,
        flags=Gtk.DialogFlags.MODAL,
        type=Gtk.MessageType.ERROR,
        message_format=message,
        buttons=(Gtk.STOCK_OK, Gtk.ResponseType.OK))
    popup_dialog.run()
    popup_dialog.destroy()
    return

def popup_yesno(parent_window, message):
    popup_dialog = Gtk.MessageDialog(
        parent=parent_window,
        flags=Gtk.DialogFlags.MODAL,
        type=Gtk.MessageType.ERROR,
        message_format=message,
        buttons=(Gtk.STOCK_NO, Gtk.ResponseType.NO, Gtk.STOCK_YES, Gtk.ResponseType.YES))
    answer = popup_dialog.run()
    popup_dialog.destroy()
    return answer