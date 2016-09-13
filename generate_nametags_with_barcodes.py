#!/usr/bin/python

#
# generate_nametags_with_barcodes.py
#       Copyright (C) 2016 Sandeep M
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
#       pylabels, a Python library to create PDFs for printing labels.
#       Copyright (C) 2012, 2013, 2014 Blair Bonnett
#
#       ReportLab open-source PDF Toolkit
#       (C) Copyright ReportLab Europe Ltd. 2000-2015
#
#       openpyxl, a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
#
# generate_nametags_with_barcodes.py is free software: 
# you can redistribute it and/or modify it under the terms of the 
# GNU General Public License as published by the Free Software Foundation, 
# either version 3 of the License, or (at your option) any later version.
#
# generate_nametags_with_barcodes.py is distributed in the hope that it 
# will be useful, but WITHOUT ANY # WARRANTY; without even the implied 
# warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
#

# ok, here we go:
from reportlab.graphics import renderPDF
from reportlab.graphics import shapes
from reportlab.graphics.barcode import code39, code128, code93
from reportlab.graphics.barcode import eanbc, qr, usps
from reportlab.graphics.shapes import Drawing 
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm, inch
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
from pprint import pprint

# for utils
from collections import OrderedDict
import re
 
#----------------------------------------------------------------------
# Create a page based on Avery 5160:  
#    portrait (8.5" X 11") sheets with 3 columns and 10 rows of labels. 
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
    # units = mm !
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
        
        #background_filename=background_filename,

    )
    return specs

#----------------------------------------------------------------------
# adjust fontsize down until it fits a width/height limit
# should really range for value instead of timidly crepping towards target
#----------------------------------------------------------------------

def fit_text_in_area(the_text,font_name,text_width_limit,text_height_limit):
    font_size = text_height_limit
    text_width = stringWidth(the_text, font_name, font_size)
    while ((text_width > text_width_limit) or (font_size > text_height_limit)):
        font_size *= 0.95
        text_width = stringWidth(the_text, font_name, font_size)

    s = shapes.String(0, 0, the_text, fontName=font_name, fontSize=font_size, textAnchor="start")
    #pprint("text_height_limit = " + str(text_height_limit))
    #pprint(s.dumpProperties())
    #pprint(s)
    return s

#----------------------------------------------------------------------
# generate strings of family name from line data
#----------------------------------------------------------------------
def get_labels_from_data (data):

    # special pattern to produce blank barcodes
    pattern_to_blank = "Zzzzzzz"

    #print("write_data")
    #pprint(data)

    # section1: the actual barcode
    num1 = data['parent_id_for_sticker'][0]
    #if (num1 > 10000): num1 -= 10000  # DEBUG
    # WORKAROUND FOR BUG: the id sometimes has a 0.5 at the end because of the way records were split
    #num1 = int(num1) + 1

    # section2: family name
    str1 = data['child_last_name'][0]
    if (pattern_to_blank in str1): str1 = " "

    # section3: parent names with & as joiner
    str2 = conjunction(data['parent_first_name'])
    if (pattern_to_blank in str2): str2 = " "

    # section4: child's names
    str3 = conjunction(data['child_first_name'])
    if (pattern_to_blank in str3): str3 = " "

    # section 4 : label number
    #str4 = str(data['index']+1) + "/" + str(data['number_of_stickers'] )
    str4 = " "


    return (num1, str1, str2, str3, str4)

#----------------------------------------------------------------------
# http://stackoverflow.com/questions/21217846/python-join-list-of-strings-with-comma-but-with-some-conditions-code-refractor
#----------------------------------------------------------------------
def conjunction(l, threshold = 5):
    length = len(l)
    l = map(str,l)
    if   length <= 2:         return " & ".join(l)
    elif length < threshold:  return ", ".join(l[:-1]) + " & " + l[-1]
    elif length == threshold: return ", ".join(l[:-1]) + " & 1 other"
    else: return ", ".join(l[:t-1]) + " & +{} others".format(length - (t - 1))

#----------------------------------------------------------------------
# adjust str height if there are any low-hanging letters (ie decenders)
#----------------------------------------------------------------------
def get_font_height(size,str):
    pattern = re.compile(r'[gjpqy]')
    if pattern.findall(str):
       size *= 1.1

    return size
#----------------------------------------------------------------------
# Create a callback function to draw each label. 
# This will be given the ReportLab drawing object to draw on, 
# the dimensions in points, and the data to put on the nametag
#----------------------------------------------------------------------

def write_data(label, width, height, data):

    (num1, str1, str2, str3, str4) = get_labels_from_data(data)

    pad = 10;   

    # section 1 : barcode
    D = Drawing(width,height)
    d = createBarcodeDrawing('Code128', value=num1,  barHeight=0.4*inch, humanReadable=True, quiet=False)
    #d = createBarcodeDrawing('I2of5', value=the_num,  barHeight=10*mm, humanReadable=True)

    barcode_width  = d.width
    barcode_height = d.height

    #d.rotate(-90)
    #d.translate( - barcode_height ,pad) # translate 

    d.translate( width-barcode_width-pad/2.0 ,0) # translate 

    #pprint(d.dumpProperties())

    #D.add(d)
    #label.add(D)
    label.add(d)


    rect = shapes.Rect(0, pad, barcode_width + pad, barcode_height+pad)
    rect.fillColor = None
    rect.strokeColor = random.choice((colors.blue, colors.red, colors.green))
    #rect.strokeWidth = d.borderStrokeWidth
    #label.add(rect)


    # section 2 : room number
    #the_text = "gr" + str(data['youngest_child_grade']) + " rm" + str(data['youngest_child_room'])
    #label.add(shapes.String(15, height-15, the_text, fontName="Judson Bold", fontSize=8, textAnchor="start"))

    # section2: family name
    # Measure the width of the name and shrink the font size until it fits.
    font_name = "Judson Bold"
    font_name = "PermanentMarker"

    # Measure the width of the name and shrink the font size until it fits.
    # try out 2 options and select the one that gives a taller font
    text_width_limit  = width  - barcode_width - pad
    text_height_limit = height / 2.0;
    s1 = fit_text_in_area(str1,font_name,text_width_limit,text_height_limit)

    text_width_limit  = width - pad
    text_height_limit = height - barcode_height
    s2 = fit_text_in_area(str1,font_name,text_width_limit,text_height_limit)

    if (s1.fontSize >= s2.fontSize): s = s1
    else:                            s = s2

    s.x = pad/2.0
    s.y = height - s.fontSize + pad / 2.0
    s.textAnchor = "start"
    label.add(s)

    family_name_height = get_font_height(s.fontSize,str1)
    family_name_width  = stringWidth(str1,font_name,s.fontSize)

    # section3: parent names
    text_width_limit  = width  - barcode_width - 2 * pad
    text_height_limit = (height - family_name_height)/2.0
    font_name = "Judson Bold"

    s = fit_text_in_area(str2,font_name,text_width_limit,text_height_limit)
    s.x = pad/2.0
    s.y = height - family_name_height - s.fontSize + pad/2.0
    s.textAnchor = "start"
    label.add(s)

    parent_name_height = get_font_height(s.fontSize,str2)

    # section4: child's names
    text_width_limit  = width  - barcode_width - 2 * pad
    text_height_limit = height - family_name_height - parent_name_height 
    font_name = "Judson Bold"

    s = fit_text_in_area(str3,font_name,text_width_limit,text_height_limit)
    s.x = pad/2.0
    s.y = height - family_name_height - parent_name_height - s.fontSize + pad/2.0
    s.textAnchor = "start"
    label.add(s)

    child_name_height = s.fontSize


    # section 4 : label number
    font_name = "Judson Bold"
    font_size = 5
    s = shapes.String(width, height - font_size, str4, fontName=font_name, fontSize=font_size, textAnchor="end")
    #s.x = width
    #s.y = 0
    #s.textAnchor = "start"
    #label.add(s)

    # section 5 : logo
    s = shapes.Image(0, 0, 25, 25, "logo.jpg")
    s.x = width - barcode_width - (barcode_width-25)/2.0 + 1
    s.y = height - pad - 15
    # enough space?
    if ((width - family_name_width - pad) > barcode_width): 
        label.add(s)


    # section 6 : "anon" label for WHP
    #if (num1 == 6710):
    #   s = shapes.Image(0, 0, 57, 34, "whp-logo.png")
    #   s.x = barcode_width + pad/2.0
    #   s.y = pad 
    #   #label.add(s)

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

def process_one_record(tag,k,v):


    if (len(v) >= 11):
        LABELS = """
        grade
        child_last_name
        child_first_name
        parent_last_name
        parent_first_name
        parent_id_for_sticker
        phone
        email
        teacher
        room
        number_of_stickers
        """


        labels=LABELS.split()
        line_item = dict(zip(labels,v))
        print("one item: len = ", len(v) )
        pprint(line_item)

        id = line_item['parent_id_for_sticker']
        # only store lines with valid id
        if (not is_number(id)):
            return
        if (id == 0):
            return


        items = {}
        if (tag.get(id) is None):
            for key in labels: 
                items[key] = [line_item[key]]
        else:
            old_items = tag[id]
            for key in labels: 
                items[key] = old_items[key] + [line_item[key]]
                
        tag[id] = items

    return

#----------------------------------------------------------------------
#
# slurp in the excel file and return a dict for easy processing
#
#----------------------------------------------------------------------

def print_one_tag(items):

    sheet.add_label(items)


#       # only print record with > 0 number of stickers
#       # otherwise, print a minimum of 3 labels
#       # align number of stickers to be easily cut, ie, multiples of 3
#
#
#        # see http://stackoverflow.com/questions/9810391/round-to-the-nearest-500-python
#       line_item['number_of_stickers']  = 1 # DEBUG OVERRIDE
#       c = 3 # number of columns 
#       x = line_item.get('number_of_stickers') 
#       if (is_number(x) and (x != 0)):
#           if (x < c ): x = c
#           else:        x = x + (c - x) % c
#           line_item['number_of_stickers'] = x
#
#
#----------------------------------------------------------------------
#
# slurp in the excel file and return a dict for easy processing
#
#----------------------------------------------------------------------
def load_records_from_excel(data_file, sheet_name):
    # load excel file--hardcoded name of workbook
    wb = load_workbook(filename=data_file, read_only=True)
    ws = wb[sheet_name] 

    # now store this in a dict with row number as the key
    records = {}
    for row in ws.rows:
        index = tuple( cell.row for cell in row)[3] # pick one col which definately has a value 
        records[index] = tuple( cell.value for cell in row)

    return records 

#----------------------------------------------------------------------
#
# process records using a helper for each one
# collect multiple records that share parent id and produce 1-to-1
# map with labels 
#
#----------------------------------------------------------------------
def process_records (records):
    tag = {}
    record_limit = 1e6   # useful for testing and runaway bugs
    count = 0

    for k,v in records.items():
        process_one_record(tag,k,v)
        count += 1
        if (count >= record_limit):
            break
        
    print("processed " ,  count, " records ")

    return tag

#----------------------------------------------------------------------
#
# check tag id compaction
#
#----------------------------------------------------------------------
def fix_tags (tag):
    count = 0

    # remove multiple values by ordered uniq list, see:
    # http://stackoverflow.com/questions/480214/how-do-you-remove-duplicates-from-a-list-in-python-whilst-preserving-order
    for id,items in tag.items():
        for k,v in items.items():
            items[k] = list(OrderedDict.fromkeys(v))
        tag[id] = items

    # can we limit id to some threshold?
    limit = 10000
    for id,items in tag.items():
        if (id > limit):
           id_short = id - limit
           if (tag.get(id_short) is not None):
                print("fix_tags: CLASH for ", limit, " for id= ", id)

    # validation tests
    for id,items in tag.items():
        if (len(items['child_last_name']) > 1):
            print("fix_tags: entry " , id, " has children with different last name" , items['child_last_name'])
 
    for id,items in tag.items():
        if (len(items['parent_first_name']) == 0):
            print("fix_tags: entry " , id, " has no parents")


    # children sometimes have names like "beiber gomez" which combine
    # parent names.  we try to catch cases where the child last name is 
    # completely different from single parent


    for id,items in tag.items():
        if (len(items['parent_first_name']) == 1):
            parent_first_name = items['parent_first_name'][0]
            parent_last_name = items['parent_last_name'][0]
            child_last_name  = items['child_last_name'][0]
            if ((parent_last_name not in child_last_name) and 
                 child_last_name not in parent_last_name):
                print("fix_tags: entry " , id, " has 1 parent with different last name from child: parent_last_name = " , parent_last_name, " child_last_name = ", child_last_name)
                # append single parent last name into first
                parent_first_last_name = parent_first_name + " " + parent_last_name
                items['parent_first_name'] = [parent_first_last_name]
                print("fix_tags: new parent name is ", parent_first_last_name)
                tag[id] = items

    return tag

#----------------------------------------------------------------------
#
# print by columns
# print out process records (ie tags) sorted by first child's last name
#
#----------------------------------------------------------------------
def print_tags_by_column (tag):
    count = 0
    
    # sorting is complex because some last names have multiple words
    # and we want to sort by num of stickers then names
    sorted_items_list = tag.values()
    sorted_items_list = sorted(sorted_items_list, key=lambda items: 
                items['child_last_name'][0].split()[-1]
    )
    sorted_items_list = sorted(sorted_items_list, key=lambda items: 
                items['number_of_stickers']
    )

    # duplicate entries for tags requiring extra column of stickers
    # OVERRIDE number_of_stickers!
    # convert number_of_stickers from 15,30,60 to 20,40,80 to make it easier
    # to cut: each column of stickers is 10
    duplicate_items_list = []
    for items in sorted_items_list:
        number_of_stickers = max(items['number_of_stickers'])
        number_of_column = 1 + int((number_of_stickers) / 15.0)
        #if (number_of_column < 1): number_of_column = 1
        for i in range(number_of_column):
            duplicate_items_list.append(items)


    # output stickers 3 at a time across for 10 in a column
    number_of_stickers_per_column = 10
    number_of_stickers_per_row    =  3
    for i in range(0,len(duplicate_items_list),number_of_stickers_per_row):
        for j in range(number_of_stickers_per_column):
            for k in range(number_of_stickers_per_row):
                index = i+k
                if (index >= len(duplicate_items_list)):
                    index  = len(duplicate_items_list) - 1
                print_one_tag(duplicate_items_list[index])
                count += 1

    print("printed " ,  count, " stickers")
    return count

#----------------------------------------------------------------------
#
# print by row
# print out process records (ie tags) sorted by first child's last name
#
#----------------------------------------------------------------------
def print_tags_by_row (tag):
    number_of_stickers_per_column = 10
    number_of_stickers_per_row    =  3
    number_of_stickers_per_page   =  number_of_stickers_per_column * number_of_stickers_per_row
    
    # sorting is complex because some last names have multiple words
    sorted_items_list = sorted(tag.values(), key=lambda
                items: items['child_last_name'][0].split()[-1])

    # duplicate entries for tags requiring extra column of stickers
    # OVERRIDE number_of_stickers!
    # convert number_of_stickers from 15,30,60 to 20,40,80 to make it easier
    # to cut: each column of stickers is 10
    duplicate_items_list = []
    for items in sorted_items_list:
        number_of_stickers = max(items['number_of_stickers'])
        number_of_stickers = number_of_stickers_per_row * int(number_of_stickers / number_of_stickers_per_row)
        if (number_of_stickers < number_of_stickers_per_row): number_of_stickers = number_of_stickers_per_row
        for i in range(number_of_stickers):
            duplicate_items_list.append(items)


    # output stickers 3 at a time across for 10 in a column
    sticker_count = 0
    for items in duplicate_items_list:
        print_one_tag(items)
        sticker_count += 1

    page_count = sticker_count / number_of_stickers_per_page
    
    print("printed " ,  sticker_count, " stickers in ", page_count , " pages")
    return sticker_count


#----------------------------------------------------------------------
#
# single tags for debug
#
#----------------------------------------------------------------------
def debug_print_tags (tag):
    count = 0
    
    sorted_items_list = sorted(tag.values(), key=lambda
        items: items['child_last_name'][0].split()[-1])

    for items in sorted_items_list:
        print_one_tag(items)
        count += 1

    print("printed " ,  count, " stickers")
    return count

#----------------------------------------------------------------------
#
# main
#
#----------------------------------------------------------------------

# single barcode per person or actually follow number_of_barcodes
DEBUG_PRINT  = 0
PRINT_BY_ROW = 0

# register some fonts, assumed to be in the same dir as this script
base_path = os.path.dirname(__file__)
font_path = os.path.join(base_path, "fonts")
registerFont(TTFont('Judson Bold',      os.path.join(font_path, 'Judson-Bold.ttf')))
registerFont(TTFont('KatamotzIkasi',    os.path.join(font_path, 'KatamotzIkasi.ttf')))
registerFont(TTFont('Magnus Cederholm', os.path.join(font_path, 'FFF_Tusj.ttf')))
registerFont(TTFont('PermanentMarker',  os.path.join(font_path, 'PermanentMarker.ttf')))


# load excel and loop through rows
data_file = 'Fallfest Barcode File.xlsx'
sheet_name = 'Barcodes'


# parse data and create
records = load_records_from_excel(data_file, sheet_name)
tag = process_records(records)
tag = fix_tags(tag)

# create the sheet with callback function write_data to process each record
specs = createAvery5160Spec()
sheet = labels.Sheet(specs, write_data, border=True)

if (DEBUG_PRINT): 
        debug_print_tags(tag)
else:
    if (PRINT_BY_ROW): 
        print_tags_by_row(tag)
    else:
        print_tags_by_column(tag)
    #endif
#endif

# save results in pdf
sheet.save('nametags.pdf')
print("{0:d} label(s) output on {1:d} page(s).".format(sheet.label_count, sheet.page_count))


