#!/usr/bin/python

#
# generate_nametags_with_barcodes.py
# 
# every year an elementary school in california runs a festival where families 
# sign up for parties and events, as well as bid for auctions and donations.
# each family is issued some stickers with unique barcode to make it easier 
# to sign up.
#
# i couldn't figure out how to get avery on-line mailmerge to do all i wanted
# (scale fonts to fit, conditionally print parent's names, repeat labels etc)
# so here we are.
#

# uses:
# 	pylabels, a Python library to create PDFs for printing labels.
# 	Copyright (C) 2012, 2013, 2014 Blair Bonnett
#
#       ReportLab open-source PDF Toolkit
#       (C) Copyright ReportLab Europe Ltd. 2000-2015
#
#       openpyxl, a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
#
# generate_nametags_with_barcodes.py is free software: you can redistribute it and/or 
# modify it under the terms of the GNU General Public License as published by 
# the Free Software Foundation, either version 3 of the License, or (at your 
# option) any later version.
#
# generate_nametags_with_barcodes.py is distributed in the hope that it will be useful, 
# but WITHOUT ANY # WARRANTY; without even the implied warranty of MERCHANTABILITY 
# or FITNESS FOR # A PARTICULAR PURPOSE.  
#

# ok, here we go:
from reportlab.graphics import renderPDF
from reportlab.graphics import shapes
from reportlab.graphics.barcode import code39, code128, code93
from reportlab.graphics.barcode import eanbc, qr, usps
from reportlab.graphics.shapes import Drawing 
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import registerFont, stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import getCodes, getCodeNames, createBarcodeDrawing

import labels
import os.path
import random
random.seed(187459)

# for excel reading
from openpyxl import load_workbook
import pprint
 
#----------------------------------------------------------------------
# Create a page based on Avery 5160:  
#    portrait (210mm x 297mm) sheets with 3 columns and 10 rows of labels. 
#
#----------------------------------------------------------------------
def createAvery5160Spec():

    f = 25.4 # conversion factor from inch to mm

    # Compulsory arguments.
    sheet_width  =  8.5 * f
    sheet_height = 11.0 * f
    columns =  3
    rows    = 10
    label_width  = 2.63 * f
    label_height = 1.00 * f
    
    # Optional arguments; missing ones will be computed later.
    left_margin   = 0.19 * f
    column_gap    = 0.12 * f
    right_margin  = 0
    top_margin    = 0.50 * f
    row_gap       = 0
    bottom_margin = 0
    
    # Optional arguments with default values.
    left_padding   = 1
    right_padding  = 1
    top_padding    = 1
    bottom_padding = 1
    corner_radius  = 2
    padding_radius = 0

    background_filename="bg.png"



    #specs = labels.Specification(210, 297, 3, 8, 65, 25, corner_radius=2)
    specs = labels.Specification(
	sheet_width, sheet_height,
	columns, rows,
	label_width, label_height,

	left_margin    = left_margin    ,
	column_gap     = column_gap     ,
	# right_margin   = right_margin   ,
	top_margin     = top_margin     ,
	row_gap        = row_gap        ,
	# bottom_margin  = bottom_margin  ,

	left_padding   = left_padding   ,
	right_padding  = right_padding  ,
	top_padding    = top_padding    ,
	bottom_padding = bottom_padding ,
	corner_radius  = corner_radius  ,
	padding_radius = padding_radius ,
	
	background_filename=background_filename,

    )
    return specs


#----------------------------------------------------------------------
# Create a function to draw each label. This will be given the ReportLab drawing
# object to draw on, the dimensions in points, and the data to put on the nametag
#----------------------------------------------------------------------
def write_data(label, width, height, data):

    #print("write_data")
    #pprint.pprint(data)

    # section 1 : barcode
    the_num = data['parent_id_for_sticker']
    d = createBarcodeDrawing('Code128', value=the_num,  barHeight=10*mm, humanReadable=True)
    #pprint.pprint(d.dumpProperties())
    barcode_width = d.width - 10
    label.add(d)

    # section 2 : room number
    the_text = "gr" + str(data['youngest_child_grade']) + " rm" + str(data['youngest_child_room'])
    label.add(shapes.String(15, height-15, the_text, fontName="Judson Bold", fontSize=8, textAnchor="start"))

    # section3: parent names
    name1 = data['firstname_parentguardian1']
    name2 = data['firstname_parentguardian2']

    # test for blank conditions
    test = (u'(blank)' in name1 or name1.isspace() , u'(blank)' in name2 or name2.isspace())
    if   (cmp(test,(True,  True )) == 0): the_text = " "
    elif (cmp(test,(False, True )) == 0): the_text = name1 + " &"
    elif (cmp(test,(True , False)) == 0): the_text = name2 + " &"
    elif (cmp(test,(False, False)) == 0): the_text = name1 + ", " + name2 + " &"

    # Measure the width of the name and shrink the font size until it fits.
    the_text = the_text.title()
    font_size = 30
    text_width = width - barcode_width
    name_width = stringWidth(the_text, "Judson Bold", font_size)
    while name_width > text_width:
        font_size *= 0.95
        name_width = stringWidth(the_text, "Judson Bold", font_size)


    label.add(shapes.String(width-2, height-20, the_text, fontName="Judson Bold", fontSize=font_size, textAnchor="end"))

    # section4: child's full name
    the_text = data['youngest_child_first_name'] + " " + data['familyname']
    the_text = the_text.title()
    # Measure the width of the name and shrink the font size until it fits.
    font_size = 100
    text_width = width - barcode_width
    name_width = stringWidth(the_text, "KatamotzIkasi", font_size)
    while name_width > text_width:
        font_size *= 0.95
        name_width = stringWidth(the_text, "KatamotzIkasi", font_size)

    # Write out the name in the centre of the label with a random colour.
    s = shapes.String( barcode_width, 20, the_text)
    s.fontName = "KatamotzIkasi"
    s.fontSize = font_size
    #s.fillColor = random.choice((colors.blue, colors.red, colors.green))
    label.add(s)

    # section 4 : label number
    the_text = str(data['index']+1) + "/" + str(data['number_of_stickers'])
    s = shapes.String(width-2, 5, the_text, fontName="Judson Bold", fontSize=6, textAnchor="end")
    label.add(s)

    # section 5 : logo
    #s = shapes.Image(barcode_width + 0.4 * (width - barcode_width), 0, 15, 15, "logo.jpg")
    #label.add(s)

#----------------------------------------------------------------------
# helper to catch blank fields in excel file
#----------------------------------------------------------------------

def is_number(s):

    if (s is None):
	return False

    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False


#----------------------------------------------------------------------
#
# create a dict from excel row, assuming all the headers match up order below
#
#----------------------------------------------------------------------

def process_one_record(k,v):
    # only process if we read 16 columns of data
    if (len(v) == 16):
	LABELS = """
	empty
	empty
	parent_id_for_sticker
	familyname
	number_of_stickers
	studentrecords_parentid
	familyname
	firstname_parentguardian1
	lastname_parentguardian1
	firstname_parentguardian2
	lastname_parentguardian2
	email_guardian_1
	email_guardian_2
	youngest_child_first_name
	youngest_child_grade
	youngest_child_room
	"""

	labels=LABELS.split()
	line_item = dict(zip(labels,v))
	#pprint.pprint(line_item)

	# only print record with > 0 number of stickers
	# otherwise, print a minimum of 3 labels
	# align number of stickers to be easily cut, ie, multiples of 3

        # see http://stackoverflow.com/questions/9810391/round-to-the-nearest-500-python
	c = 3 # number of columns 
	x = line_item.get('number_of_stickers') 
	if (is_number(x) and (x != 0)):
	    if (x < c ): x = c
	    else: 	 x = x + (c - x) % c

	    pprint.pprint(line_item)
	    for i in range(x):
		line_item['index'] = i
	        sheet.add_label(line_item)

#----------------------------------------------------------------------
#
# slurp in the excel file and return a dict for easy processing
#
#----------------------------------------------------------------------
def load_records_from_excel(data_file):
    # load excel file--hardcoded name of workbook
    wb = load_workbook(filename=data_file, read_only=True)
    ws = wb['Sticker Data'] 

    # now store this in a dict with row number as the key
    records = {}
    for row in ws.rows:
	index = tuple( cell.row for cell in row)[-1]
	records[index] = tuple( cell.value for cell in row)

    return records 


#----------------------------------------------------------------------
#
# main
#
#----------------------------------------------------------------------


# register some fonts, assumed to be in the same dir as this script
base_path = os.path.dirname(__file__)
registerFont(TTFont('Judson Bold',   os.path.join(base_path, 'Judson-Bold.ttf')))
registerFont(TTFont('KatamotzIkasi', os.path.join(base_path, 'KatamotzIkasi.ttf')))

# create the sheet
specs = createAvery5160Spec()
sheet = labels.Sheet(specs, write_data, border=True)

# load excel and loop through rows
data_file = '../sample_excel/Sample Sticker Data File.xlsx'
records = load_records_from_excel(data_file)

for k,v in records.items():
    process_one_record(k,v)

# save results
sheet.save('nametags.pdf')
print("{0:d} label(s) output on {1:d} page(s).".format(sheet.label_count, sheet.page_count))


