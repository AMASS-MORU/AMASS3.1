#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Data verification logfile reports systematically.

# Created on 20th April 2022
import pandas as pd #for creating and manipulating dataframe
from datetime import date #for generating today date
from pathlib import Path #for retrieving input's path
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_JUSTIFY #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib.colors import * #for importing color palette
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_commonlib as AL
# Check that is there either CSV or XLSX with the specific file name or not
def checkxlsorcsv(spath,sfilename) :
    bfound = False
    try:
        if Path(spath + sfilename + ".csv").is_file():
            bfound = True
        else:
            if Path(spath + sfilename + ".xlsx").is_file():
                bfound = True
    except:
        bfound = False
    return bfound
def checkpoint(str_filename):
    return Path(str_filename).is_file()

def marked_idx(df,datatype="df",row_per_page=29):
    lst = []
    lst_idx = []
    if datatype == "df":
        for i in list(df.index):
            if i %row_per_page == 0:
                lst_idx.append(i)
    elif datatype == "lst":
        for i in range(len(df)):
            if i %row_per_page == 0:
                lst_idx.append(i)
    return lst_idx

def report_title(c,title_name,pos_x,pos_y,font_color,font_size=20):
    c.setFont("Helvetica-Bold", font_size) # define a large bold Helvetica
    c.setFillColor(font_color) #define font color
    c.drawString(pos_x,pos_y,title_name)

def report_context(c,context_list,pos_x,pos_y,wide,height,font_size=10,font_align=TA_JUSTIFY,line_space=18,left_indent=0):
    context_list_style = []
    style = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=font_size,leftIndent=left_indent,alignment=font_align)
    for cont in context_list:
        cont_1 = Paragraph(cont, style)
        context_list_style.append(cont_1)
    f = Frame(pos_x,pos_y,wide,height,showBoundary=0)
    return f.addFromList(context_list_style,c)

def report_table_appendix(df):
    list_font = AL.setlocalfont()
    LOCALFONT = list_font[0]
    LOCALFONT_BOLD = list_font[1]
    #print(LOCALFONT)
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),LOCALFONT),
                           ('FONTSIZE',(0,0),(-1,-1),8),
                           ('FONTSIZE',(0,0),(0,-1),8),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

#Retrieving user value from dictionary_for_microbiology_data.xlsx
#return value is string of user value.
def retrieve_uservalue(dict_df, amass_name, col_amass="amass_name", col_user="user_name"):
    return dict_df.loc[dict_df[col_amass]==amass_name,:].reset_index().loc[0,col_user]

#Retrieving values from xxx_results.xlsx
#return value is string of that value.
def retrieve_results(df, str_find_1, col_find_1, str_find_2="", col_find_2="", col_res="Values"):
    result = ""
    if str_find_2 == "":
        result = df.loc[df[col_find_1]==str_find_1,col_res].tolist()[0]
    else:
        result = df.loc[(df[col_find_1]==str_find_1)&(df[col_find_2]==str_find_2),col_res].tolist()[0]
    return result

#Assigning Not available of NA
#Return value: assigned variables
def assign_na_toinfo(str_info, coverpage=False):
    if str_info == "empty001_micro" or str_info == "empty001_hosp" or str_info == "NA" or str_info == "" or str_info != str_info:
        if coverpage:
            str_info = "Not available"
        else:
            str_info = "NA"
    else:
        str_info = str(str_info)
    return str_info

def prepare_unicode(str_unicode):
    result = ""
    if len(str(str_unicode)) > 40:
        result = str(str_unicode)[:40]
    else:
        result = str(str_unicode)
    return result

# def prepare_unicode(str_unicode):
#     result = ""
#     if len(str(str_unicode)) > 50:
#         if len(str(str_unicode)) < 100:
#             result = str(str_unicode)[:50] + "\n" + str(str_unicode)[50:]
#         elif len(str(str_unicode)) >= 100 and len(str(str_unicode)) < 150:
#             result = str(str_unicode)[:50] + "\n" + str(str_unicode)[50:100] + "\n" + str(str_unicode)[100:]
#         elif len(str(str_unicode)) >= 150 and len(str(str_unicode)) < 200:
#             result = str(str_unicode)[:50] + "\n" + str(str_unicode)[50:100] + "\n" + str(str_unicode)[100:150] + "\n" + str(str_unicode)[150:]
#         else:
#             result = str(str_unicode)[:50] + "\n" + str(str_unicode)[50:100] + "\n" + str(str_unicode)[100:150] + "\n" + str(str_unicode)[150:200] + "\n" + str(str_unicode)[200:]
#     else:
#         result = str(str_unicode)
#     return result
