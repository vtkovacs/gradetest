'''
Created on Dec 15, 2015
make a pdf of the test scores
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
from reportlab import platypus
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus.tables import Table
from reportlab.lib import colors
from gi.repository import Gtk
import shared
from exifread import process_file
import re
import os
pdf_filename = None
csv_filename = None
supporting_info = ""
name_store = None
do_pdf = True

def save():
    if do_pdf :
        savepdf()
    else:
        savecsv()   
    return
    
def savepdf():
    if os.path.exists(pdf_filename):
        overwrite = shared.popup_yesno(shared.mainwindow, pdf_filename + ' exists, overwrite?')
        if overwrite != Gtk.ResponseType.YES:
            return        
    supporting_info_print = re.sub('\n', '<br/>', supporting_info)
    doc = platypus.SimpleDocTemplate(pdf_filename)
    content = [platypus.Spacer(1,0.1*inch)]
    style = getSampleStyleSheet()["Normal"]
    content.append(platypus.Paragraph(supporting_info_print, style))
    content.append(platypus.Spacer(1,0.2*inch))
    content.append(platypus.Paragraph('Template: ' + shared.test_template, style))
    content.append(platypus.Paragraph('Scanfile: ' + shared.test_scanfile + ' - Date scanned: ' + str(shared.scan_date), style))
    content.append(platypus.Spacer(1,0.2*inch))
#     header_table_data = [[platypus.Paragraph(supporting_info_print, style)]]
#     header_table_data.append([platypus.Paragraph('Template: ' + shared.test_template, style)])
#     header_table_data.append([platypus.Paragraph('Scanfile: ' + shared.test_scanfile + ' - Date scanned: ' + str(scan_date), style)])
#     headerTable = Table(header_table_data)
#     content.append(headerTable)
    
    scores_table_data = [['Name', 'Score', 'Quality', 'Sheet', 'Verified']]
    for row in name_store:
        [name, score, quality, sheet, verified] = row
        if verified:
            verified_print = 'yes'
        else:
            verified_print = ''
        scores_table_data.append([name, score, quality, str(sheet), verified_print])
    
    scoresTable = Table(scores_table_data, style = [('GRID',(0,0),(-1,-1),1,colors.gray)])
    content.append(scoresTable)
    doc.build(content)
    return

def savecsv():
    if os.path.exists(csv_filename):
        overwrite = shared.popup_yesno(shared.mainwindow, csv_filename + ' exists, overwrite?')
        if overwrite != Gtk.ResponseType.YES:
            return
    supporting_info_print =  supporting_info
    doc = open(csv_filename,'w')
    doc.write('"' + supporting_info_print + '"\n')
    doc.write('Template:, ' + shared.test_template +'\n')
    doc.write('Scanfile:, ' + shared.test_scanfile + ',Date scanned: ,' + str(shared.scan_date) + '\n')
    doc.write('\nName,Score,Quality,Sheet,Verified\n')
    
    for row in name_store:
        [name, score, quality, sheet, verified] = row
        if verified:
            verified_print = 'yes'
        else:
            verified_print = ''
            
        doc.write(name + ',' + str(score) + ',' + quality + ',' + str(sheet) + ',' + verified_print +'\n')
    
    return