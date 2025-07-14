#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Supplementary data indicators reports systematically.

# Created on 20th April 2022
import pandas as pd #for creating and manipulating dataframe
import datetime #for formatting date
from pathlib import Path #for retrieving input's path
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib.colors import * #for importing color palette
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables



#Checking process is either able for running or not
#df_config: Dataframe of config file
#str_process_name: process name in string fmt
#return value: Boolean; True when parameter is set "yes", False when parameter is set "no"
def check_config(df_config, str_process_name):
    config_lst = df_config.iloc[:,0].tolist()
    result = ""
    if df_config.loc[config_lst.index(str_process_name),"Setting parameters"] == "yes":
        result = True
    else:
        result = False
    return result

#Checking file available
#return : boolean ; True when file is available; False when file is not available
def checkpoint(str_filename):
    return Path(str_filename).is_file()

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

#Retrieving list of user's values
#return value: list of user's values
def retrieve_userlist(df_dict, amass_name, col_amass="amass_name", col_user="user_name"):
    lst = df_dict.loc[df_dict[col_amass]==amass_name,:].reset_index().loc[:,col_user].tolist()
    return [i for i in lst if i != ""]

#Preparing tables for summary page i-iii
def create_table_summary_v3(df,col_1,col_2,col_3,col_4):
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
    df[col_3] = df[col_3].replace(regex=["Proportion of notifiable antibiotic-pathogen combinations"],value="Proportion of notifiable\nantibiotic-pathogen combinations")
    df[col_3] = df[col_3].replace(regex=["Proportion of potential errors in the AST results"],value="Proportion of potential errors\nin the AST results")
    for idx in df.index: 
        #Assigning NA to ""
        if df.loc[idx,col_4] == "":
            df.at[idx,col_4] = "NA"
        else:
            pass
        #Assigning * and **
        note = ""
        if idx in [0,1,2,3]:
            note = "*"
        else:
            note = "**"
        if idx in [0,4,8]:
            df.loc[idx,col_1] = Paragraph("<b>" + df.loc[idx,col_1] + "</b>",style_summary)
            df.loc[idx,col_2] = Paragraph("<b>" + str(df.loc[idx,col_2]) + "</b>",style_summary)
            df.loc[idx,col_3] = Paragraph("<b>" + df.loc[idx,col_3] + "</b>",style_summary)
            df.loc[idx,col_4] = Paragraph("<b>" + df.loc[idx,col_4] + note + "</b>",style_summary)
        else:
            df.loc[idx,col_4] = df.loc[idx,col_4] + note
            #Assigning italic
            if df.loc[idx,col_1] != "":
                df.loc[idx,col_1] = Paragraph(prepare_org_core_v2(df.loc[idx,col_1]),style_summary)
            else:
                pass

    df = df.loc[df[col_1]!="",:].rename(columns={col_1:"",col_2:"Indicators",col_3:"Description of indicators",col_4:"Number of\nblood samples\n(%)"})
    df_col = [list(df.columns)]
    df = df.values.tolist()
    df = df_col + df
    return df

#Preparing tables for summary page i-iii
def create_table_summary_v2(df,col_1,col_2,col_3,col_4):
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
    df[col_3] = df[col_3].replace(regex=["Proportion of notifiable antibiotic-pathogen combinations"],value="Proportion of notifiable\nantibiotic-pathogen combinations")
    df[col_3] = df[col_3].replace(regex=["Proportion of potential errors in the AST results"],value="Proportion of potential errors\nin the AST results")
    for idx in df.index: 
        #Assigning NA to ""
        if df.loc[idx,col_4] == "":
            df.at[idx,col_4] = "NA"
        else:
            pass
        #Assigning * and **
        note = ""
        if idx in [0,1,2,3]:
            note = "*"
        else:
            note = "**"
        if idx in [0,4,8]:
            df.loc[idx,col_1] = Paragraph("<b>" + df.loc[idx,col_1] + "</b>",style_summary)
            df.loc[idx,col_2] = Paragraph("<b>" + str(df.loc[idx,col_2]) + "</b>",style_summary)
            df.loc[idx,col_3] = Paragraph("<b>" + df.loc[idx,col_3] + "</b>",style_summary)
            df.loc[idx,col_4] = Paragraph("<b>" + df.loc[idx,col_4] + note + "</b>",style_summary)
        else:
            df.loc[idx,col_4] = df.loc[idx,col_4] + note
        #
    df = df.loc[df[col_1]!="",:].rename(columns={col_1:"",col_2:"Indicators",col_3:"Description of indicators",col_4:"Number of\nblood samples\n(%)"})
    df_col = [list(df.columns)]
    df = df.values.tolist()
    df = df_col + df
    return df

def prepare_org_core_v2(str_org, text_line=1, text_style="full", text_work="table", text_work_drug="N", text_bold="N"):
    lst_org = str_org.split(" ")
    ##Preparing capital character for organism
    if len(lst_org) <= 2: #['MRSA'] or ['Brucella', 'spp'] >>> pass
        pass
    else:                 
        lst_org[-2] = lst_org[-2].capitalize()
            
    ##Preparing format of organism
    if text_style == "full":
        pass
    elif text_style == "short":
        if len(lst_org) == 2: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                pass
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0][0] + ".", lst_org[-1]]
        elif len(lst_org) == 3: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                pass
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0], lst_org[1][0] + ".", lst_org[-1]]

    ##Preparing organism for applying to work
    for i in range(len(lst_org)):
        if "spp" in lst_org[i] or "spp." in lst_org[i] or "other" in lst_org[i].lower() or "contaminants" in lst_org[i].lower():
            pass
        else:
            if text_work == "table":
                if text_work_drug == "N":  #S. pneumoniae (No durg information)
                    if "serogroup" in lst_org[i].lower() or len(lst_org[i]) < 3:
                        pass    #<i>Salmonella</i> serogroup C
                    else:
                        lst_org[i] = "<i>" + lst_org[i] + "</i>"    #<i>S.</i> <i>pneumoniae</i>
                else:                   #Penicillin-NS S. pneumoniae
                    if i == 0:          #1st index >>> Penicillin-NS (passed)
                        pass
                    else:               #other index >>> <i>S.</i> <i>pneumoniae</i> (italic added)
                        lst_org[i] = "<i>" + lst_org[i] + "</i>"
            elif text_work == "graph":  
                if text_work_drug == "N":  #S. pneumoniae (No durg information)
                    lst_org[i] = "$" + lst_org[i] + "$"         #$S.$ $pneumoniae$
                else:                   #Penicillin-NS S. pneumoniae
                    if i == 0:          #1st index >>> Penicillin-NS (passed)
                        pass
                    else:               #other index >>> $S.$ $pneumoniae$ (italic added)
                        lst_org[i] = "$" + lst_org[i] + "$"

    ##Preparing organism in line
    if len(lst_org) > 2:
        if text_line == 1: #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', 'Salmonella', 'spp']
            lst_org[-2] = lst_org[-2]
        else:             #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', '\nSalmonella', 'spp']
            lst_org[-2] = "\n" + lst_org[-2]

    ##Preparing bold organism (updated on 01/11/21)
    if text_bold=="Y":
        lst_org[0] = "<b>" + lst_org[0]
        lst_org[-1] = lst_org[-1] + "</b>"
    else:
        pass
    return (" ".join(lst_org))

def prepare_org_annexA(str_org, text_line=1, text_style="full", text_work="table"):
    #str_org : Organism with string datatype ex. 'Non-typhoidal salmonella spp'
    #text_line : Number of line for setting organism name format (1--dafault, or 2) ex. 'Non-typhoidal Salmonella spp' or 'Non-typhoidal \nSalmonella spp'
    #text_style : Format type of organism name ("full"--default, or "short") ex. 'Burkholderia pseudomallei or 'B. pseudomallei'
    #text_work : Type of work that organism will be applied ("table"--default, or "graph")
    lst_org = str_org.split(" ")
    ##Preparing capital character for organism
    if len(lst_org) <= 2: #['MRSA'] or ['Brucella', 'spp'] >>> pass
        pass
    else:                 
        lst_org[-2] = lst_org[-2].capitalize()    
    ##Preparing format of organism
    if text_style == "full":
        pass
    elif text_style == "short":
        if len(lst_org) == 2: 
            if "spp" == lst_org[-1] or "spp." == lst_org[-1]: #['Brucella', 'spp'] >>> pass
                lst_org = [lst_org[0], "spp."]
            else:                                             #['Burkholderia', 'pseudomallei'] >>> ['B.', 'pseudomallei']
                lst_org = [lst_org[0][0] + ".", lst_org[-1]] 
    ##Preparing organism for applying to work
    for i in range(len(lst_org)):
        if "spp" in lst_org[i] or "spp." in lst_org[i]:
            lst_org[i] = "spp."
        else:
            if text_work == "table":
                lst_org[i] = "<i>" + lst_org[i] + "</i>"
            elif text_work == "graph":
                lst_org[i] = "$" + lst_org[i] + "$"    
    ##Preparing organism in line
    if len(lst_org) > 2:
        if text_line == 1: #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', 'Salmonella', 'spp']
            lst_org[-2] = lst_org[-2]
        else:             #['Non-typhoidal', 'salmonella', 'spp'] >>> ['Non-typhoidal', '\nSalmonella', 'spp']
            lst_org[-2] = "\n" + lst_org[-2]
    return (" ".join(lst_org))

#default : 30 rows per page
def create_lst_marked_page(df, row_per_page=30):
    i = 0
    index_mark_page = [i]
    while i + row_per_page < len(df.index):
        if i == 0:
            i = i + row_per_page - 1
        else:
            i = i + row_per_page
        index_mark_page.append(i)
    return index_mark_page

#Escherichia coli isolate that is resistant to Cephems
#(Ertapenem=R, Imipenem=R, Meropenem=R, Cefepime=R, Cefotaxime=R, Ceftazidime=R, Ceftriaxone=R)
#to
#Escherichia coli isolate that is resistant to Cephems
#(Ertapenem=R, Imipenem=R, Meropenem=R, Cefepime=R, 
#Cefotaxime=R, Ceftazidime=R, Ceftriaxone=R)
def reformat_alerted_message(str_alert):
    result = ""
    res = [i.start() for i in re.finditer(", ", str_alert)]
    if len(res) > 4:
        if len(res) <= 6:
            result = str_alert[:res[3]] + "," + "\n" + str_alert[res[3]+2:]
        else:
            result = str_alert[:res[3]] + "," + "\n" + str_alert[res[3]+2:res[6]] + "," + "\n" + str_alert[res[6]+2:]
    else:
        result = str_alert
    return result

#Reformatting except organism for summary table (Potential blood contamination)
#Return value: string of (formatted) except organism
def prepare_except_org_for_summarytable(str_except_org):
    temp_lst = str_except_org.split(",")
    temp_lst_1 = [prepare_org_core_v2(temp_lst[i].capitalize()) for i in range(len(temp_lst))]
    result = ""
    if len(temp_lst_1) > 1:
        result = ", ".join(temp_lst_1[:len(temp_lst_1)-1]) + ", and " + temp_lst_1[-1]  #<i>Staphylococcus</i> <i>aureus</i>, and <i>Staphylococcus</i> <i>ludunensis</i>
    else:
        result = ", ".join(temp_lst_1) #<i>Bacillus</i> <i>anthracis</i>
    return result

def prepare_summarytable(df_summary):
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=8,alignment=TA_LEFT)
    df_summary = df_summary.reset_index().drop(columns=["index"])
    for idx in df_summary.index:
        if df_summary.at[idx,"rule_organism"].lower() == "all":
            df_summary.at[idx,"rule_organism"] = Paragraph(df_summary.loc[idx,"rule_organism"].capitalize(), style_summary)
        else:
            df_summary.at[idx,"rule_organism"] = Paragraph(prepare_org_core_v2(df_summary.loc[idx,"rule_organism"].capitalize()).replace("<i>Burkholderia</i> <i>Cepacia</i> <i>complex</i>","<i>Burkholderia cepacia</i> complex").replace("<i>Enterobacter</i> <i>Cloacae</i> <i>complex</i>","<i>Enterobacter cloacae</i> complex").replace("<i>Enterobacteriaceae</i>","Enterobacteriaceae"), style_summary)
    return df_summary

def prepare_list_alerted_selectrow(df, col_indicator, col_status):
    return df.loc[(df[col_indicator]!="") & (df[col_status]=="yes"),].fillna("").reset_index().drop(columns=["index"])

def prepare_list_alerted_formatorg(df, col_org="mapped_sci", col_oth=""):
    font_style = ParagraphStyle('normal',fontName='Helvetica',fontSize=7,alignment=TA_CENTER)
    for idx in df.index:
        org = df.loc[idx,col_org]
        org_fmt = prepare_org_core_v2(df.loc[idx,col_org].capitalize())
        df.at[idx,col_org] = Paragraph(org_fmt,font_style) #reformatting organism name
        if col_oth != "":
            df.at[idx,col_oth] = reformat_alerted_message(df.loc[idx,col_oth].replace(org,org_fmt)) #reformatting organism name
            df.at[idx,col_oth] = Paragraph(df.loc[idx,col_oth], font_style)
        else:
            pass
    return df

def prepare_list_alerted_selectcol(df, col_status, col_hn, col_spcdate, col_spctype="mapped_spctype", col_org="mapped_sci", col_oth="", col_refmt_oth=""):
    if col_oth != "":
        df_1 = df.loc[:,[col_hn,col_spcdate,col_spctype,col_org,col_oth]].rename(columns={col_hn:"Hospital\nnumber",col_spcdate:"Specimen\ncollection\ndate",col_spctype:"Specimen\ntype",col_org:"Organism",col_oth:col_refmt_oth})
    else:
        df_1 = df.loc[:,[col_hn,col_spcdate,col_spctype,col_org]].rename(columns={col_hn:"Hospital\nnumber",col_spcdate:"Specimen\ncollection\ndate",col_spctype:"Specimen\ntype",col_org:"Organism"})
    return df_1

#Formating date (i.e. specimen collection date) into "01 Jan 2012" format
#Return value: dataframe with formated date column
def format_date_forexportation(df,col_date):
    df["spcdate_fmt"] = ""
    for idx in df.index:
        if isinstance(df.loc[idx,col_date], str):
            df.at[idx,"spcdate_fmt"]=df.loc[idx,col_date]
        else:
            print (df.loc[idx,col_date], isinstance(df.loc[idx,col_date], str), df.loc[idx,col_date].strftime("%d %b %Y"))
            df.at[idx,"spcdate_fmt"]=df.loc[idx,col_date].strftime("%d %b %Y")
    return df

#Correcting page number based on marked_index_page
def correct_content_page(list_content_page):
    list_content_page_correct = []
    for x in range(len(list_content_page)):
        if x != len(list_content_page)-1: #If x is not 1st index and is not lasted index >>> append summary value to list
            list_content_page_correct.append(sum(list_content_page[:x+1]))
    return list_content_page_correct

def correct_content_page_v2(lst_page):
    lst_page_correct = []
    for x in range(len(lst_page)):
        z = 0
        if x == 0:
            z = lst_page[x]
        else:
            z = 1 + sum(lst_page[:x])
        lst_page_correct.append(z)
    return lst_page_correct

def report_title(c,title_name,pos_x,pos_y,font_color,font_size=20):
    c.setFont("Helvetica-Bold", font_size) # define a large bold Helvetica
    c.setFillColor(font_color) #define font color
    c.drawString(pos_x,pos_y,title_name)

def report_context(c,context_list,pos_x,pos_y,wide,height,font_size=10,font_align=TA_JUSTIFY,line_space=18,left_indent=0):
    context_list_style = []
    style = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=font_size,leftIndent=left_indent,alignment=font_align)
    style_content = ParagraphStyle('normal',fontName='Helvetica',leading=line_space,fontSize=0.5,leftIndent=left_indent,alignment=font_align)
    for cont in context_list:
        if "." == cont:
            cont_1 = Paragraph(cont, style_content)
        else:
            cont_1 = Paragraph(cont, style)
        context_list_style.append(cont_1)
    f = Frame(pos_x,pos_y,wide,height,showBoundary=0)
    return f.addFromList(context_list_style,c)

def report_todaypage(c,pos_x,pos_y,footer_information):
    c.setFont("Helvetica", 9) # define a large bold font
    c.setFillColor('#3e4444')
    c.drawString(pos_x,pos_y,footer_information)

def report_table_summary(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                        #    ('FONT',(0,1),(0,-1),'Helvetica-Oblique'),
                           ('FONTSIZE',(0,0),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,1),(-1,-1),'LEFT'),
                           ('ALIGN',(0,0),(-1,0),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=[2.2*inch,0.8*inch,2.2*inch,1.3*inch])

def report_table_summary_v2(df):
    return Table(df,style=[('FONT',(0,0),(0,-1),'Helvetica-Bold'),
                           ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                           ('FONT',(1,1),(1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),8),
                           ('FONTSIZE',(0,0),(0,-1),8),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=[2.0*inch,2.1*inch,2.0*inch])

def report_table_summary_1_v2(df):
    return Table(df,style=[('FONT',(0,0),(0,-1),'Helvetica-Bold'),
                           ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                           ('FONT',(1,1),(1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),8),
                           ('FONTSIZE',(0,0),(0,-1),8),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=[4.1*inch,2.0*inch])

def report_table_appendix(df):
    return Table(df,style=[('FONT',(0,0),(0,-1),'Helvetica-Bold'),
                           ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                           ('FONT',(1,1),(1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),7),
                           ('FONTSIZE',(0,0),(0,-1),8),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=[0.7*inch,0.7*inch,0.7*inch,1.5*inch,2.8*inch])

def report_table_appendix_2(df):
    return Table(df,style=[('FONT',(0,0),(0,-1),'Helvetica-Bold'),
                           ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
                           ('FONT',(1,1),(1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),7),
                           ('FONTSIZE',(0,0),(0,-1),8),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.black),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], colWidths=[1.5*inch,1.5*inch,1.5*inch,1.5*inch])