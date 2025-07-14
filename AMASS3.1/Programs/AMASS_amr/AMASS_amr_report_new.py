#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.1 (AMASS version 3.1) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.

# Created on 20 APR 2022
# Last update on: 15 JUL 2025 #3.1 3104
import math as m
import gc 
import psutil, os
import logging #for creating error_log
import re #for manipulating regular expression
import pandas as pd #for creating and manipulating dataframe
import matplotlib.pyplot as plt #for creating graph (pyplot)
import matplotlib #for importing graph elements
from pathlib import Path #for retrieving input's path
from PIL import Image #for importing image
import seaborn as sns #for creating graph
from datetime import date #for generating today date
from datetime import datetime 
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_LEFT, TA_CENTER #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
#from AMASS_amr_commonlib_report import * #for importing amr functions
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
import AMASS_amr_commonlib_report as ARC

#Const
path_result = AC.CONST_PATH_RESULT
path_input = AC.CONST_PATH_ROOT 
sec1_res_i = AC.CONST_FILENAME_sec1_res_i
sec1_num_i = AC.CONST_FILENAME_sec1_num_i
sec2_res_i = AC.CONST_FILENAME_sec2_res_i
sec2_amr_i = AC.CONST_FILENAME_sec2_amr_i
sec2_org_i = AC.CONST_FILENAME_sec2_org_i
sec2_pat_i = AC.CONST_FILENAME_sec2_pat_i
sec3_res_i = AC.CONST_FILENAME_sec3_res_i
sec3_amr_i = AC.CONST_FILENAME_sec3_amr_i
sec3_pat_i = AC.CONST_FILENAME_sec3_pat_i
sec4_res_i = AC.CONST_FILENAME_sec4_res_i
sec4_blo_i = AC.CONST_FILENAME_sec4_blo_i
sec4_pri_i = AC.CONST_FILENAME_sec4_pri_i
sec5_res_i = AC.CONST_FILENAME_sec5_res_i
sec5_com_i = AC.CONST_FILENAME_sec5_com_i
sec5_hos_i = AC.CONST_FILENAME_sec5_hos_i
sec5_com_amr_i = AC.CONST_FILENAME_sec5_com_amr_i
sec5_hos_amr_i = AC.CONST_FILENAME_sec5_hos_amr_i
sec6_res_i = AC.CONST_FILENAME_sec6_res_i
sec6_mor_byorg_i = AC.CONST_FILENAME_sec6_mor_byorg_i
sec6_mor_i = AC.CONST_FILENAME_sec6_mor_i
secA_res_i = AC.CONST_FILENAME_secA_res_i
secA_pat_i = AC.CONST_FILENAME_secA_pat_i
secA_mor_i = AC.CONST_FILENAME_secA_mor_i
secA_res_i_A11 = AC.CONST_FILENAME_secA_res_i_A11
secA_pat_i_A11 = AC.CONST_FILENAME_secA_pat_i_A11
secB_blo_i = AC.CONST_FILENAME_secB_blo_i
secB_blo_mon_i = AC.CONST_FILENAME_secB_blo_mon_i
##paragraph variable
iden1_op = "<para leftindent=\"35\">"
iden2_op = "<para leftindent=\"70\">"
iden3_op = "<para leftindent=\"105\">"
iden_ed = "</para>"
bold_blue_ital_op = "<b><i><font color=\"#000080\">"
bold_blue_ital_ed = "</font></i></b>"
bold_blue_op = "<b><font color=\"#000080\">"
bold_blue_ed = "</font></b>"
green_op = "<font color=darkgreen>"
green_ed = "</font>"
add_blankline = "<br/>"
tab1st = "&nbsp;"
tab4th = "&nbsp;&nbsp;&nbsp;&nbsp;"

#Global variables/CONST for report 

def sub_printprocmem(sstate,logger) :
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at state " +sstate + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
    except:
        AL.printlog("Error get process memory usage at " + sstate,True,logger)

def canvas_printpage(c,curpage,lastpage,today=date.today().strftime("%d %b %Y"),bisrotate90=False,ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1,istartpage=1):
    ARC.report_todaypage(c,55,30,"Created on: "+today)
    imode = ipagemode
    if imode == 1:
        ARC.report_todaypage(c,270,30,"Page " + str(curpage) + " of " + str(lastpage))
    elif imode == 2:
        ARC.report_todaypage(c,250,30,ssectionanme + " Page " + str(curpage) + " of " + str(isecmaxpage))
    else:
        iseccurpage = curpage - istartpage + 1
        if iseccurpage <= isecmaxpage:
            ARC.report_todaypage(c,270,30,"Page " + str(curpage) + " of " + str(lastpage))
        else:
            iasc = 65 + iseccurpage - isecmaxpage - 1
            curpage = istartpage + isecmaxpage -1
            ARC.report_todaypage(c,270,30,"Page " + str(curpage) + "-" + str(chr(iasc)) +" of " + str(lastpage))
    if bisrotate90 == True:
        c.rotate(90)
    c.showPage()
    

def cover(c,logger,today=date.today().strftime("%d %b %Y")):
    sec1_file = AC.CONST_PATH_RESULT + sec1_res_i
    section1_result = pd.DataFrame()
    try:
        section1_result = pd.read_csv(sec1_file).fillna("NA")
    except Exception as e:
        AL.printlog("Error read file : " + sec1_file + " : " + str(e),True,logger)
        return 
    if len(section1_result) <= 0:
        AL.printlog("Error no record in file : " + sec1_file ,True,logger)
        return
    ##paragraph variables
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    add_blankline = "<br/>"
    ##variables
    if len(section1_result) > 0:
        hospital_name   = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Hospital_name","Values"].tolist()[0]), coverpage=True)
        country_name    = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Country","Values"].tolist()[0]), coverpage=True)
        contact_person  = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_person","Values"].tolist()[0]), coverpage=True)
        contact_address = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_address","Values"].tolist()[0]), coverpage=True)
        contact_email   = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Contact_email","Values"].tolist()[0]), coverpage=True)
        notes   = str(section1_result.loc[section1_result["Parameters"]=="notes_on_the_cover","Values"].tolist()[0])
        #change in v3.0 build 3010 to use overall min/max date
        spc_date_start  = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=True)
        spc_date_end    = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=True)
        #spc_date_start  = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=True)
        #spc_date_end    = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=True)
    else:
        hospital_name    = "NA"
        country_name    = "NA"
        contact_person  = "NA"
        contact_address = "NA"
        contact_email   = "NA"
        notes           = "NA"
        spc_date_start  = "NA"
        spc_date_end    = "NA"

    ##text
    cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
    cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
    cover_1_3 = "<b>Data from:</b>"
    cover_1_4 = bold_blue_op + str(spc_date_start) + " to " + str(spc_date_end) + bold_blue_ed
    cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
    cover_2_1 = "<b>Contact person:</b>  " + bold_blue_op + contact_person + bold_blue_ed
    cover_2_2 = "<b>Contact address:</b>  " + bold_blue_op + contact_address + bold_blue_ed
    cover_2_3 = "<b>Contact email:</b>  " + bold_blue_op + contact_email + bold_blue_ed
    cover_2_4 = "<b>Generated on:</b>  " + bold_blue_op + today + bold_blue_ed
    cover_2_5 = "<b>Software version:</b>  " + bold_blue_op + AC.CONST_SOFTWARE_VERSION + bold_blue_ed
    cover_2 = [cover_2_1,cover_2_2,cover_2_3,cover_2_4,cover_2_5]
    if notes == "" or notes == "NA":
        pass
    else:
        cover_2 = cover_2 + ["<b>Notes:</b>  " + notes]
    ########### COVER PAGE ############
    c.setFillColor('#FCBB42')
    c.rect(0,590,800,20, fill=True, stroke=False)
    c.setFillColor(forestgreen)
    c.rect(0,420,800,150, fill=True, stroke=False)
    ARC.report_title(c,'Antimicrobial Resistance (AMR)',0.7*inch, 515,'white',font_size=28)
    ARC.report_title(c,'Surveillance report',0.7*inch, 455,'white',font_size=28)
    ARC.report_context(c,cover_1, 0.7*inch, 2.2*inch, 460, 240, font_size=18,line_space=26)
    ARC.report_context(c,cover_2, 0.7*inch, 0.5*inch, 460, 140, font_size=11)
    c.showPage()
def generatedby(c,logger):
    sec1_file = AC.CONST_PATH_RESULT + sec1_res_i
    section1_result = pd.DataFrame()
    try:
        section1_result = pd.read_csv(sec1_file).fillna("NA")
    except Exception as e:
        AL.printlog("Error read file : " + sec1_file + " : " + str(e),True,logger)
        return 
    if len(section1_result) <= 0:
        AL.printlog("Error no record in file : " + sec1_file ,True,logger)
        return
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    add_blankline = "<br/>"
    ##variables
    if len(section1_result) > 0:
        hospital_name   = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Hospital_name","Values"].tolist()[0]), coverpage=False)
        country_name    = ARC.assign_na_toinfo(str_info=str(section1_result.loc[section1_result["Parameters"]=="Country","Values"].tolist()[0]), coverpage=False)
        spc_date_start  = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0]), coverpage=False)
        spc_date_end    = ARC.assign_na_toinfo(str_info=str(section1_result.loc[(section1_result["Type_of_data_file"]=="overall_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0]), coverpage=False)
    else:
        hospital_name    = "NA"
        country_name    = "NA"
        spc_date_start  = "NA"
        spc_date_end    = "NA"

    generatedby_1_1  = "Generated by"
    generatedby_1_2  = "AutoMated tool for Antimicrobial resistance Surveillance System (AMASS) version 3.1"
    generatedby_1_3  = "(released on " + AC.CONST_SOFTWARE_RELEASE +")"
    generatedby_1_5  = "AMASS application is available under the Creative Commons Attribution 4.0 International Public License (CC BY 4.0). The application can be downloaded at : <u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>"
    generatedby_1_6  = "AMASS application used microbiology data and hospital admission data files that are stored in the same folder as the application (AMASS.bat) to generate this report."
    generatedby_1_7  = "The goal of AMASS application is to enable hospitals with microbiology data available in electronic formats to analyze their own data and generate AMR surveillance reports promptly. If hospital admission date data are available, the reports will additionally be stratified by infection origin (community−origin or hospital−origin). If mortality data (such as patient discharge outcome data) are available, a report on mortality involving AMR infection will be added."
    generatedby_1_8  = "This automatically generated report has limitations, and requires users to understand those limitations and use the summary data in the report with careful interpretation."
    generatedby_1_9  = "A valid report could have local implications and much wider benefits if shared with national and international organizations."
    generatedby_1_10  = "This automatically generated report is under the jurisdiction of the hospital to copy, redistribute, and share with any individual or organization."
    generatedby_1_11 = "This automatically generated report contains no patient identifier, similar to standard reports on cumulative antimicrobial susceptibility."
    generatedby_1_12 = "For any query on AMASS, please contact:"
    generatedby_1_13 = "Chalida Rangsiwutisak (chalida@tropmedres.ac),"
    generatedby_1_14 = "Cherry Lim (cherry@tropmedres.ac), and"
    generatedby_1_15 = "Direk Limmathurotsakul (direk@tropmedres.ac)"
    generatedby_1 = ["<b>" + generatedby_1_1 + "</b>", 
                    generatedby_1_2, 
                    generatedby_1_3,  
                    iden1_op + add_blankline + generatedby_1_5 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_6 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_7 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_8 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_9 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_10 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_11 + iden_ed, 
                    iden1_op + add_blankline + generatedby_1_12 + iden_ed, 
                    iden1_op + generatedby_1_13 + iden_ed, 
                    iden1_op + generatedby_1_14 + iden_ed, 
                    iden1_op + generatedby_1_15 + iden_ed]
    generatedby_2_1 = "Suggested title for citation:"
    generatedby_2_2 = "Antimicrobial resistance surveillance report, " + hospital_name + ","
    generatedby_2_3 = country_name + ", "+ str(spc_date_start) + " to " + str(spc_date_end) + "."
    generatedby_2 = ["<b>" + generatedby_2_1 + "</b>", 
                    generatedby_2_2, 
                    generatedby_2_3]
    ########### GENERATED BY ##########
    ARC.report_context(c,generatedby_1, 1.0*inch, 0.7*inch, 460, 700, font_size=11, line_space=16)
    ARC.report_context(c,generatedby_2, 1.0*inch, 0.6*inch, 460, 80, font_size=11, line_space=16)
    c.showPage()
    
def introduction(c,logger,startpage, lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    add_blankline = "<br/>"
    ##Page1
    intro_page1_1 = "Antimicrobial resistance (AMR) is a global health crisis [1]. "+ \
                "The report by Lord Jim O'Neill estimated that 700,000 global deaths could be attributable to AMR in 2015, and projected that the annual death toll could reach 10 million by 2050 [1]. "+ \
                "However, data of AMR surveillance from low and middle−income countries (LMICs) are scarce [1,2], and data of mortality associated with AMR infections are rarely available. "+ \
                "A recent study estimated that 19,000 deaths are attributable to AMR infections in Thailand annually, using routinely available microbiological and hospital databases [3]. "+ \
                "The study also proposed that hospitals in LMICs should utilize routinely available microbiological and hospital admission databases to generate reports on AMR surveillance systematically [3]."
    intro_page1_2_1 = "Reports on AMR surveillance can have a wide range of benefits [2]; including"
    intro_page1_2_2 = "− characterization of the frequency of resistance and organisms in different facilities and regions;"
    intro_page1_2_3 = "− prospective and retrospective information on emerging public health threats;"
    intro_page1_2_4 = "− evaluation and optimization of local and national standard treatment guidelines;"
    intro_page1_2_5 = "− evaluation of the impact of interventions beyond antimicrobial guidelines that aim to reduce AMR; and"
    intro_page1_2_6 = "− data sharing with national and international organizations to support decisions on resource allocation for interventions against AMR and to inform the implementation of action plans at national and global levels."
    intro_page1_3 = "When reporting AMR surveillance results, it is generally recommended that (a) duplicate results of bacterial isolates are removed, and (b) reports are stratified by infection origin (community−origin or hospital−origin), if possible [2]. "+ \
                    "Many hospitals in LMICs lack time and resources needed to analyze the data (particularly to deduplicate data and to generate tables and figures), write the reports, and to release the data or reports [4]."
    intro_page1_4 = "AutoMated tool for Antimicrobial resistance Surveillance System (AMASS) was developed as an offline, open−access and easy−to−use application that allows a hospital to perform data analysis independently and generate AMR proportion and AMR frequency reports stratified by infection origin from routinely collected electronic databases. The application was built in a free software environment. The application has been placed within a user−friendly interface that only requires the user to double−click on the application icon. AMASS application can be downloaded at: <u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>"
    intro_page1 = [intro_page1_1, 
                add_blankline + intro_page1_2_1, 
                iden1_op + intro_page1_2_2 + iden_ed, 
                iden1_op + intro_page1_2_3 + iden_ed, 
                iden1_op + intro_page1_2_4 + iden_ed, 
                iden1_op + intro_page1_2_5 + iden_ed, 
                iden1_op + intro_page1_2_6 + iden_ed, 
                add_blankline + intro_page1_3, 
                add_blankline + intro_page1_4]
    ##Page2
    intro_page2_1_1 = "AMASS version 3.1 additionally generates reports on notifiable bacterial diseases in Annex A and on data indicators (including proportion of contaminants and discordant AST results) in Annex B for the \"microbiology data\" file that is used to generate this report. A careful review of the Annex B could help readers and data owners to identify potential errors in the microbiology data used to generate the report." #3.1 3104
    intro_page2_1_2 = "AMASS version 3.1 also separately generates Supplementary data indictors report (in PDF and Excel formats) in a new folder \"Report_with_patient_identifiers\" to support users to check and validate records with notifiable bacteria, notifiable antibiotic-pathogen combinations, infrequent phenotypes or potential errors in the AST results at the local level. The identifiers listed include hospital number and specimen collection date. The files are generated in a separate folder \"Report_with_patient_identifiers\" so that it is clear that users should not share or transfer the Supplementary Data Indictors report (in PDF and Excel format) to any party outside of the hospital without data security management and confidential agreement." #3.1 3104
    intro_page2_1 = [intro_page2_1_1, add_blankline + intro_page2_1_2]
    intro_page2_2_1 = "References:"
    intro_page2_2_2 = "[1] O'Neill J. (2014) Antimicrobial resistance: tackling a crisis for the health and wealth of nations. Review on antimicrobial resistance. http://amr−review.org. (accessed on 3 Dec 2018)."
    intro_page2_2_3 = "[2] World Health Organization (2018) Global Antimicrobial Resistance Surveillance System (GLASS) Report. Early implantation 2016−2017. http://apps.who.int/iris/bitstream/handle/10665/259744/9789241513449−eng.pdf. (accessed on 3 Dec 2018)"
    intro_page2_2_4 = "[3] Lim C., et al. (2016) Epidemiology and burden of multidrug−resistant bacterial infection in a developing country. Elife 5: e18082."
    intro_page2_2_5 = "[4] Ashley EA, Shetty N, Patel J, et al. Harnessing alternative sources of antimicrobial resistance data to support surveillance in low−resource settings. J Antimicrob Chemother. 2019; 74(3):541−546."
    intro_page2_2_6 = "[5] Clinical and Laboratory Standards Institute (CLSI). Analysis and Presentation of Cumulative Antimicrobial Susceptibility Test Data, 4th Edition. 2014. (accessed on 21 Jan 2020)"
    intro_page2_2_7 = "[6] European Antimicrobial Resistance Surveillance Network (EARS−Net). Antimicrobial resistance (AMR) reporting protocol 2018. (accessed on 21 Jan 2020)"
    intro_page2_2_8 = "[7] European Committee on Antimicrobial Susceptibility Testing (EUCAST). www.eucast.org (accessed on 21 Jan 2020)"
    intro_page2_2 = ["<b>" + intro_page2_2_1 + "</b>", 
                    intro_page2_2_2, 
                    intro_page2_2_3, 
                    intro_page2_2_4, 
                    intro_page2_2_5, 
                    intro_page2_2_6, 
                    intro_page2_2_7, 
                    intro_page2_2_8]
    ########## INTRO: PAGE1 ###########
    ARC.report_title(c,'Introduction',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
    ARC.report_context(c,intro_page1, 1.07*inch, 1.0*inch, 460, 650, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
    ########## INTRO: PAGE2 ###########
    ARC.report_context(c,intro_page2_1, 1.07*inch, 4.5*inch, 460, 450, font_size=11)
    ARC.report_context(c,intro_page2_2, 1.07*inch, 0.5*inch, 460, 270, font_size=9,font_align=TA_LEFT,line_space=14) #Reference
    u = inch/10.0
    c.setLineWidth(2)
    c.setStrokeColor(black)
    p = c.beginPath()
    p.moveTo(70,315) # start point (x,y)
    p.lineTo(7.2*inch,315) # end point (x,y)
    c.drawPath(p, stroke=1, fill=1)
    canvas_printpage(c,startpage+1,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage)
def nodata_micro(c,stitle,startpage,lastpage,today,ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    page1_1_1 = "Not applicable because microbiology data file is not available or the format of microbiology data file is not supported. Please save microbiology data file in excel format (.xlsx) or csv (.csv; UTF-8)." #3.1 3104
    page1_1 = [page1_1_1]
    ARC.report_title(c,stitle,1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
    
def section1_nodata(c,startpage,lastpage,today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    nodata_micro(c,'Section [1]: Data overview',startpage,lastpage,today)

def section2_nodata(c,startpage,lastpage,today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    nodata_micro(c,'Section [2]: AMR proportion report',startpage,lastpage,today)

def section3_nodata(c,startpage,lastpage,today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    section3_page1_1_1 = "Proportions of antimicrobial−resistance infection stratified by origin of infection are not calculated because hospital admission date data is not available and infection origin variable is not available." #3.1 3104
    section3_page1_1 = [section3_page1_1_1] 
    ######### SECTION3: PAGE1 #########
    ARC.report_title(c,'Section [3]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_context(c,section3_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
def section4_nodata(c,startpage,lastpage,today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    section4_page1_1_1 = "Incidence of infections per 100,000 tested population is not calculated because data on blood specimens with no growth is not available." #3.1 3104
    section4_page1_1 = [section4_page1_1_1]
    ######### SECTION4: PAGE1 #########
    ARC.report_title(c,'Report [4]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,section4_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
def section5_nodata(c,startpage,lastpage,today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    section5_page1_1_1 = "Incidence of infections per 100,000 tested population stratified by infection origin is not calculated because data on blood specimens with no growth is not available, or stratification by origin of infection cannot be done (due to hospital admission date variable is not available)." #3.1 3104
    section5_page1_1 = [section5_page1_1_1]
    ######### SECTION5: PAGE1 #########
    ARC.report_title(c,'Report [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_context(c,section5_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
def section6_nodata(c,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    section6_page1_1_1 = "Mortality in AMR antimicrobial-susceptible infections is not applicable because hospital admission data file is not available, or in−hospital outcome (in hospital admission data file) is not available." #3.1 3104
    section6_page1_1 = [section6_page1_1_1]  
    ######### SECTION6: PAGE1 #########
    ARC.report_title(c,'Report [6] Mortality in AMR',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_context(c,section6_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)  
def annexA_nodata(c,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    annexA_page1_1_1 = "Supplementary report on notifiable bacterial disease is not applicable because microbiology data file is not available." #3.1 3104
    annexA_page1_1 = [annexA_page1_1_1]
    ######### ANNEXA: PAGE1 #########
    ARC.report_title(c,'Annex A: Supplementary report on notifiable bacterial',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'diseases',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_context(c,annexA_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
def annexA11_nodata_mortality(c,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##Page1
    annexA_page1_1_1 = "Notifiable bacterial infections are not applicable because hospital admission data file is not available." #3.1 3104
    annexA_page1_1 = [annexA_page1_1_1]
    ######### ANNEXA: PAGE3 #########
    ARC.report_title(c,'Annex A1b: Notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16) #3.1 3104
    ARC.report_context(c,annexA_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
def annexA_nodata_mortality(c,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##Page1
    annexA_page1_1_1 = "Mortality involving the notifiable bacterial diseases is not applicable because hospital admission data file is not available, or in−hospital outcome (in hospital admission data file) is not available." #3.1 3104
    annexA_page1_1 = [annexA_page1_1_1]
    ######### ANNEXA: PAGE3 #########
    ARC.report_title(c,'Annex A2: Mortality involving notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,annexA_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
def annexB_nodata(c,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    annexB_page1_1_1 = "Supplementary report on data indicators is not applicable because microbiology data file is not available, list of indicators is not available, or number of observation is not estimated." #3.1 3104
    annexB_page1_1 = [annexB_page1_1_1]
    ######### ANNEXB: PAGE1 #########
    ARC.report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,annexB_page1_1, 1.0*inch, 8.5*inch, 460, 100, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
def section1(c,logger,bishosp_ava,section1_result, section1_table, startpage, lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0])
    spc_date_end    = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0])
    blo_num         = str(section1_result.loc[(section1_result["Type_of_data_file"]=="microbiology_data")&(section1_result["Parameters"]=="Number_of_records"),"Values"].tolist()[0])
    if bishosp_ava:
        hos_date_start = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Minimum_date"),"Values"].tolist()[0])
        hos_date_end   = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Maximum_date"),"Values"].tolist()[0])
        hos_num        = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Number_of_records"),"Values"].tolist()[0])
        patient_days   = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Patient_days"),"Values"].tolist()[0])
        patient_days_his = str(section1_result.loc[(section1_result["Type_of_data_file"]=="hospital_admission_data")&(section1_result["Parameters"]=="Patient_days_his"),"Values"].tolist()[0])
    else:
        hos_date_start = "NA"
        hos_date_end   = "NA"
        hos_num        = "NA"
        patient_days   = "NA"
        patient_days_his = "NA"

    ##Page1
    section1_page1_1_1 = "An overview of the data detected by AMASS application is generated by default. "+ \
                    "The summary is based on the raw data files saved within the same folder as the application file (AMASS.bat)."
    section1_page1_1_2 = "Please review and validate this section carefully before proceeds to the next section."
    section1_page1_1 = [section1_page1_1_1, 
                        add_blankline + section1_page1_1_2]
    section1_page1_2_1 = "The microbiology data file (stored in the same folder as the application file) had:" #3.1 3104
    section1_page1_2_2 = bold_blue_ital_op + blo_num + bold_blue_ital_ed + " specimen data records with collection dates ranging from "
    section1_page1_2_3 = bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section1_page1_2_4 = "The hospital admission data file (stored in the same folder as the application file) had:" #3.1 3104
    section1_page1_2_5 = bold_blue_ital_op + hos_num + bold_blue_ital_ed + " admission data records with hospital admission dates ranging from "
    section1_page1_2_6 = bold_blue_ital_op + hos_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + hos_date_end + bold_blue_ital_ed
    section1_page1_2_7 = "The total number of patient-days was " + bold_blue_ital_op + patient_days + bold_blue_ital_ed + "."
    section1_page1_2_8 = "The total number of patient-days at risk of BSI of hospital-origin was " + bold_blue_ital_op + patient_days_his + bold_blue_ital_ed + "."
    section1_page1_2 = [section1_page1_2_1, 
                        iden1_op + "<i>" + section1_page1_2_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section1_page1_2_3 + "</i>" + iden_ed, 
                        add_blankline + section1_page1_2_4, 
                        iden1_op + "<i>" + section1_page1_2_5 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section1_page1_2_6 + "</i>" + iden_ed,
                        iden1_op + add_blankline + section1_page1_2_7 + iden_ed,
                        iden1_op + add_blankline + section1_page1_2_8 + iden_ed]
    section1_page1_3_1 = "[1] If the periods of the data in microbiology data and hospital admission data files are not similar, " + \
                        "the automatically−generated report should be interpreted with caution. " + \
                        "AMASS generates the reports based on the available data." #3.1 3104
    section1_page1_3_2 = "[2] A patient is defined as at risk of BSI of hospital-origin when the patient is admitted to the hospital for more than two calendar days with calendar day one equal to the day of admission."
    section1_page1_3 = [green_op + section1_page1_3_1 + green_ed, 
                        green_op + section1_page1_3_2 + green_ed]
    ##Page2
    section1_page2_1_1 = "Data was stratified by month to assist detection of missing data, and verification of whether the month distribution of data records in microbiology data file and hospital admission data file reflected the microbiology culture frequency and admission rate of the hospital, respectively. " + \
                        "For example if the number of specimens in the microbiology data file reported below is lower than what is expected, please check the raw data file and data dictionary files." #3.1 3104
    section1_page2_1 = [section1_page2_1_1]
    section1_page2_2_1 = "[1] Additional general demographic data will be made available in the next version of AMASS application."
    section1_page2_2 = [green_op + section1_page2_2_1 + green_ed]
    ######### SECTION1: PAGE1 #########
    ARC.report_title(c,'Section [1]: Data overview',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section1_page1_1, 1.0*inch, 7.3*inch, 460, 150, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 7.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section1_page1_2, 1.0*inch, 4.0*inch, 460, 250, font_size=11)
    ARC.report_title(c,'Note:',1.07*inch, 3.5*inch,'darkgreen',font_size=12)
    ARC.report_context(c,section1_page1_3, 1.0*inch, 1.8*inch, 460, 120, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION1: PAGE2 #########
    ARC.report_title(c,'Reporting period by months:',1.07*inch, 10.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section1_page2_1, 1.0*inch, 8.0*inch, 460, 150, font_size=11)
    table_draw = ARC.report1_table(section1_table)
    table_draw.wrapOn(c, 600, 400)
    table_draw.drawOn(c, 1.5*inch, 4.3*inch)
    ARC.report_title(c,'Note:',1.07*inch, 3.0*inch,'darkgreen',font_size=12)
    ARC.report_context(c,section1_page2_2, 1.0*inch, 2.3*inch, 460, 50, font_size=11)
    canvas_printpage(c,startpage+1,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 

def section2(c,logger,result_table, summary_table, lst_org_format,lst_numpat, lst_org_short, list_sec2_org_table,list_sec2_atbcountperorg,list_sec2_atbnote, 
             startpage, lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden3_op = "<para leftindent=\"105\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0])
    spc_date_end    = str(result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0])
    blo_num         = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0])
    blo_num_neg     = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_negative","Values"].tolist()[0])
    blo_num_pos     = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive","Values"].tolist()[0])
    blo_num_pos_org = str(result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0])
    ##Page1
    section2_page1_1_1 = "An AMR proportion report is generated by default, even if the hospital admission data file is unavailable. "+ \
                        "This is to enable hospitals with only microbiology data available to utilize the de−duplication and report generation functions of AMASS. " + \
                        "This report is without stratification by origin of infection." #3.1 3104
    section2_page1_1_2 = "The report generated by AMASS application version 3.1 includes only blood samples. " + \
                        "The next version of AMASS will include other specimen types, including cerebrospinal fluid (CSF), urine, stool, and other specimens."
    section2_page1_1 = [section2_page1_1_1, 
                        add_blankline + section2_page1_1_2]

    section2_page1_2 = []
    for i in range(len(lst_org_format)):
        section2_page1_2.append(iden1_op + "− " + lst_org_format[i] + iden_ed)
    section2_page1_3_1 = "The microbiology data file had:" #3.1 3104
    section2_page1_3_2 = "Sample collection dates ranged from " + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + " to " + bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section2_page1_3_3 = "Number of records of blood specimens collected within the above date range:"
    section2_page1_3_4 = blo_num + " blood specimen records" #3.1 3104
    section2_page1_3_5 = "Number of records of blood specimens with *negative culture (no growth):"
    section2_page1_3_6 = blo_num_neg + " blood specimen records" #3.1 3104
    section2_page1_3_7 = "Number of records of blood specimens with culture positive for a microorganism:"
    section2_page1_3_8 = blo_num_pos + " blood specimen records" #3.1 3104
    section2_page1_3_9 = "Number of records of blood specimens with culture positive for organism under this surveillance:"
    section2_page1_3_10 = blo_num_pos_org + " blood specimen records" #3.1 3104
    section2_page1_3 = [section2_page1_3_1, 
                        iden1_op + "<i>"             + section2_page1_3_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>"             + section2_page1_3_3 + "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section2_page1_3_4 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>"             + section2_page1_3_5 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section2_page1_3_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>"             + section2_page1_3_7 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section2_page1_3_8 + bold_blue_ital_ed + iden_ed, 
                        iden3_op + "<i>"             + section2_page1_3_9 + "</i>" + iden_ed, 
                        iden3_op + bold_blue_ital_op + section2_page1_3_10 + bold_blue_ital_ed + iden_ed]
    ##Page2
    section2_page2_1_1 = "AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period as described in the method. " + \
                        "The number of patients with positive samples is as follows:"
    section2_page2_2_1 = "*The negative culture included data values specified as \"no growth\" in the dictionary file for microbiology data (details on data dictionary files are in the method section) to represent specimens with negative culture for any microorganism. " #3.1 3104
    section2_page2_2_2 = "**Only the first isolate for each patient per specimen type, per pathogen, and per evaluation period was included in the analysis."
    section2_page2_3_1 = "The following figures and tables show the proportion of patients with blood culture positive for antimicrobial non−susceptible isolates."
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section2_page2_3_1 = "The following figures and tables show the proportion of patients with blood culture positive for antimicrobial resistant isolates."
    section2_page2_1 = [section2_page2_1_1]
    section2_page2_2 = [green_op + section2_page2_2_1 + green_ed, 
                        green_op + section2_page2_2_2 + green_ed]
    section2_page2_3 = [section2_page2_3_1]
    ##Page3-7
    section2_note_1 = "*Proportion of non−susceptible (NS) isolates represents the number of patients with blood culture positive for non−susceptible isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                    "AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that AST unknown results are more than 30% of the total number of patients with blood culture positive for the organism. "
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section2_note_1 = "*Proportion of R represents the number of patients with blood culture positive for resistant isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                        "AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that AST unknown results are more than 30% of the total number of patients with blood culture positive for the organism. "

    section2_note_2_1 = "CI=confidence interval; NA=not available/reported/tested; "

    ######### SECTION2: PAGE1 #########
    ARC.report_title(c,'Section [2]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section2_page1_1, 1.0*inch, 6.5*inch, 460, 200, font_size=11)
    ARC.report_title(c,'Organisms under this surveillance:',1.07*inch, 6.8*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section2_page1_2, 1.0*inch, 4.0*inch, 460, 200, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 4.0*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section2_page1_3, 1.0*inch, 0.2*inch, 460, 270, font_size=11, font_align=TA_LEFT)
    canvas_printpage(c,startpage,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION2: PAGE2 #########
    ARC.report_context(c,section2_page2_1, 1.0*inch, 9*inch, 460, 100, font_size=11)
    ARC.report_context(c,section2_page2_2, 1.0*inch, 3.3*inch, 460, 120, font_size=11)
    ARC.report_context(c,section2_page2_3, 1.0*inch, 1.8*inch, 460, 50, font_size=11)
    table_draw = ARC.report2_table(summary_table)
    
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 5.5*inch)
    
    canvas_printpage(c,startpage+1,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION2: PAGE3-7 #########
    bIsStartofnewpage = True
    bPrintPage = False
    iPage = startpage+1
    iPos_head_y = 0
    iPos_fig_y = 0
    iPos_tbl_y = 0
    list_head_y =[9.3,5.2]
    list_fig_y =[6.18,2.08]
    list_tbl_y =[9.98,5.90]
    for i in range(len(lst_org_format)):
        try:
            """
            Layout for pic
            0.3*inch is left margin
            6.3*inch is the page 1st organism offset from bottom page to bottom of picture (increase will move image up)
            2.25*inch is the page 2st organism offset from bottom (increase will move image up)
            3.5*inch is normal high for image of 1st and 2st org, if height is differ adjust 6.3*inch and 2.25 inch acording to the different
            if only 1 org per page using the 6.3*inch and adjust by different of image high and 3.5*inch normal high
            """
            ifig_dis_H = ARC.cal_sec2and3_fig_height(list_sec2_atbcountperorg[i])
            ifig_dis_offset = 3.5 - ifig_dis_H
            #Print position calculation
            bPrintPage = False
            if list_sec2_atbcountperorg[i] > AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
                #Fullpage
                #iPos_head_y = 9.4
                #iPos_fig_y = 6.3+ifig_dis_offset
                #iPos_tbl_y =((6.45+3.5)-((len(list_sec2_org_table[i])*0.25)+0.5))
                iPos_head_y = list_head_y[0]
                iPos_fig_y = list_fig_y[0]+ifig_dis_offset
                iPos_tbl_y =((list_tbl_y[0])-((len(list_sec2_org_table[i])*0.25)+0.5))
                list_curnote = list_sec2_atbnote[i]
                bPrintPage = True
                bIsStartofnewpage = True
            elif bIsStartofnewpage: 
                #Half top page
                #iPos_head_y = 9.4
                #iPos_fig_y = 6.3+ifig_dis_offset
                #iPos_tbl_y =((6.45+3.5)-((len(list_sec2_org_table[i])*0.25)+0.5))
                iPos_head_y = list_head_y[0]
                iPos_fig_y = list_fig_y[0]+ifig_dis_offset
                iPos_tbl_y =((list_tbl_y[0])-((len(list_sec2_org_table[i])*0.25)+0.5))
                list_curnote = list_sec2_atbnote[i]
                #Check if next require fullpage then print page (note + page no)
                ii = i+1
                if ii>= len(lst_org_format):
                    bPrintPage = True
                else:
                    if list_sec2_atbcountperorg[ii] <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
                        bPrintPage = False
                    else:
                        bPrintPage = True
                bIsStartofnewpage = False
            else:
                #Half bottom page  
                #iPos_head_y = 5.4
                #iPos_fig_y = 2.25+ifig_dis_offset
                #iPos_tbl_y =((2.45+3.5)-((len(list_sec2_org_table[i])*0.25)+0.5))
                iPos_head_y = list_head_y[1]
                iPos_fig_y = list_fig_y[1]+ifig_dis_offset
                iPos_tbl_y =((list_tbl_y[1])-((len(list_sec2_org_table[i])*0.25)+0.5))
                list_curnote = list_curnote+list_sec2_atbnote[i]
                list_curnote = list(set(list_curnote))
                bPrintPage = True
                bIsStartofnewpage = True
            #Print content
            ARC.report_title(c,'Section [2]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
            line_1_1 = ["<b>" + "Blood: " + lst_org_format[i] + "</b>"]
            line_1_2 = [bold_blue_ital_op + "( No. of patients = " + str(lst_numpat[i]) + " )" + bold_blue_ital_ed]
            ARC.report_context(c,line_1_1, 1.0*inch, iPos_head_y*inch, 300,50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_2, 5.2*inch, iPos_head_y*inch, 200,50, font_size=12, font_align=TA_LEFT)
            c.drawImage(path_result + 'Report2_AMR_' + lst_org_short[i] +".png", 0.3*inch, iPos_fig_y*inch, preserveAspectRatio=True, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
            table_draw = ARC.report2_table_nons(list_sec2_org_table[i])
            table_draw.wrapOn(c,500, 300)
            table_draw.drawOn(c, 4.0*inch, iPos_tbl_y*inch)
            #Print page (note + page no) if needed
            if bPrintPage:
                scurnote = section2_note_1 + section2_note_2_1 + ARC.get_atbnote(list_curnote)
                ARC.report_context(c,[scurnote] ,1.0*inch,0.4*inch,460,120, font_size=9,line_space=12)
                iPage = iPage + 1
                canvas_printpage(c,iPage,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
                bIsStartofnewpage = True
        except Exception as e:
             AL.printlog("Error SECTION 2 at " + lst_org_format[i] +  " : " + str(e),True,logger) 
             logger.exception(e)
             pass

def section3(c,logger,sec3_res, sec3_pat_val,lst_org_format, lst_numpat_CO,lst_numpat_HO, lst_org_short,list_sec3_org_table_CO,list_sec3_org_table_HO,list_sec3_atbcountperorg,list_sec3_atbnote_CO,list_sec3_atbnote_HO,
             startpage,  lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = str(sec3_res.loc[sec3_res["Parameters"]=="Minimum_date","Values"].tolist()[0])
    spc_date_end    = str(sec3_res.loc[sec3_res["Parameters"]=="Maximum_date","Values"].tolist()[0])
    pat_num_pos_org     = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0])
    pat_num_pos_org_com = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_community_origin_BSI","Values"].tolist()[0])
    pat_num_pos_org_hos = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_hospital_origin_BSI","Values"].tolist()[0])
    pat_num_pos_org_unk = str(sec3_res.loc[sec3_res["Parameters"]=="Number_of_patients_with_unknown_origin_BSI","Values"].tolist()[0])
    ##Page1
    section3_page1_1_1 = "An AMR proportion report with stratification by origin of infection is generated only if admission date data are available in the raw data file(s) with the appropriate specification in the data dictionaries."
    section3_page1_1_2 = "Stratification by origin of infection is used as a proxy to define where the bloodstream infection (BSI) was contracted (hospital versus community)."
    section3_page1_1_3 = "The definitions of infection origin proposed by the WHO GLASS are used. In brief, community−origin BSI is defined as patients in the hospital for less than or equal to two calendar days when the first specimen culture postive for the pathogen was taken. " + \
                        "Hospital−origin BSI is defined as patients admitted for more than two calendar days when the first specimen culture positive for the pathogen was taken."
    section3_page1_1 = [section3_page1_1_1, 
                        add_blankline + section3_page1_1_2, 
                        add_blankline + section3_page1_1_3]
    section3_page1_2_1 = "The data included in the analysis to generate the report had:"
    section3_page1_2_2 = "<i>" + "Sample collection dates ranged from " + "</i>" + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        "<i>" + " to " + "</i>" +\
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section3_page1_2_3 = "*Number of patients with blood culture positive for pathogen under the surveillance:"
    section3_page1_2_4 = pat_num_pos_org + " patients"
    section3_page1_2_5 = "**Number of patients with community−origin BSI:"
    section3_page1_2_6 = pat_num_pos_org_com + " patients"
    section3_page1_2_7 = "**Number of patients with hospital−origin BSI:"
    section3_page1_2_8 = pat_num_pos_org_hos + " patients"
    section3_page1_2_9 = "***Number of patients with unknown infection of origin status:"
    section3_page1_2_10 = pat_num_pos_org_unk + " patients"
    section3_page1_2 = [section3_page1_2_1, 
                        iden1_op + section3_page1_2_2 + iden_ed, 
                        iden1_op +"<i>" + section3_page1_2_3 + "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section3_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_5 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_7 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_8 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + "<i>" + section3_page1_2_9 + "</i>" + iden_ed, 
                        iden2_op + bold_blue_ital_op + section3_page1_2_10 + bold_blue_ital_ed + iden_ed]
    ##Page2
    section3_page2_1_1 = "NA=not applicable (hospital admission date or infection origin data are not available)"
    section3_page2_1_2 = "*Only the first isolate for each patient per specimen type per pathogen under the reporting period is included in the analysis. Please refer to Section [2] for details on how this number was calculated from the raw microbiology data file." #3.1 3104
    section3_page2_1_3 = "**The definitions of infection origin proposed by the WHO GLASS is used. In brief, community−origin BSI was defined as patients in the hospital for less than or equal to two calendar days when the first blood culture positive for the pathogen was taken."
    section3_page2_1_4 = "Hospital−origin BSI was defined as patients admitted for more than two calendar days when the first specimen culture positive for the pathogen was taken."
    section3_page2_1_5 = "Please refer to the \"Methods\" section for more details on the definitions used." #3.1 3104
    section3_page2_1_6 = "***Unknown origin could be because admission date data are not available or the patient was not hospitalised."
    section3_page2_1 = [green_op + section3_page2_1_1 + green_ed, 
                        green_op + add_blankline + section3_page2_1_2 + green_ed, 
                        green_op + section3_page2_1_3 + green_ed, 
                        green_op + section3_page2_1_4 + green_ed, 
                        green_op + section3_page2_1_5 + green_ed, 
                        green_op + section3_page2_1_6 + green_ed]
    section3_page2_2_1 = "The following figures and tables below show the proportion of patients with blood culture positive for antimicrobial non−susceptible isolates stratified by infection of origin."
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section3_page2_2_1 = "The following figures and tables below show the proportion of patients with blood culture positive for antimicrobial resistant isolates stratified by infection of origin."
    section3_page2_2 = [section3_page2_2_1]
    section3_note_1 = "*Proportion of non−susceptible (NS) isolates represents the number of patients with blood culture positive for non−susceptible isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                    "AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that AST unknown results are more than 30% of the total number of patients with blood culture positive for the organism. "
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section3_note_1 = "*Proportion of R represents the number of patients with blood culture positive for resistant isolates (numerator) over the total number of patients with blood culture positive for the organism and the organism was tested for susceptibility against the antibiotic (denominator). " + \
                    "AMASS application de−duplicated the data by including only the first isolate per patient per specimen type per evaluation period. Grey bars indicate that AST unknown results are more than 30% of the total number of patients with blood culture positive for the organism. "
    section3_note_2_1 = "CI=confidence interval; NA=not available/reported/tested"
    ######### SECTION3: PAGE1 #########
    ARC.report_title(c,'Section [3]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section3_page1_1, 1.0*inch, 5.9*inch, 460, 250, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 5.7*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section3_page1_2, 1.0*inch, 2.8*inch, 460, 200, font_size=11)
    canvas_printpage(c, startpage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION3: PAGE2 #########
    ARC.report_title(c,'Note',1.07*inch, 6.0*inch,'darkgreen',font_size=12)
    ARC.report_context(c,section3_page2_1, 1.0*inch, 2.5*inch, 460, 250, font_size=11)
    ARC.report_context(c,section3_page2_2, 1.0*inch, 1.5*inch, 460, 50, font_size=11)
    table_draw = ARC.report3_table(sec3_pat_val)
    table_draw.wrapOn(c, 500, 300)
    #table_draw.drawOn(c, 3.2*inch, 6.9*inch)
    table_draw.drawOn(c, 1.0*inch, 6.9*inch)
    canvas_printpage(c, startpage+1, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION3: PAGE3-12 #########   
    ipage = startpage+2
    for i in range(len(lst_org_format)):
        try:
            bIsOrgfitonepage = True
            if list_sec3_atbcountperorg[i] <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
                bIsOrgfitonepage = False
            ifig_dis_H = ARC.cal_sec2and3_fig_height(list_sec3_atbcountperorg[i]) 
            sOrgI = "_" + str(i)
            #print y position [y is bIsOrgfitonepage,y is not bIsOrgfitonepage]
            #list_head_y =[9.4,5.4]
            #list_fig_y =[6.3,2.25]
            #list_tbl_y =[6.45,2.45]
            list_head_y =[9.3,5.2]
            list_fig_y =[6.18,2.08]
            list_tbl_y =[9.98,5.90]
            """
            Layout for pic
            0.3*inch is left margin
            6.3*inch is the page 1st organism offset from bottom page to bottom of picture (increase will move image up)
            2.25*inch is the page 2st organism offset from bottom (increase will move image up)
            3.5*inch is normal high for image of 1st and 2st org, if height is differ adjust 6.3*inch and 2.25 inch acording to the different
            if only 1 org per page using the 6.3*inch and adjust by different of image high and 3.5*inch normal high
            """
            #CO
            ARC.report_title(c,'Section [3]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
            ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
            line_1_1 = ["<b>" + "Blood: " + lst_org_format[i] + "</b>"]
            line_1_2 = ["<b>" + "Community-origin" + "</b>"]
            line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(lst_numpat_CO[i]) + " )" + bold_blue_ital_ed]
            ARC.report_context(c,line_1_1, 1.0*inch, list_head_y[0]*inch, 300, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_2, 4.0*inch, list_head_y[0]*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_3, 5.5*inch, list_head_y[0]*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            c.drawImage(path_result + 'Report3_AMR_' + lst_org_short[i] + "_Community.png", 0.3*inch, (list_fig_y[0]+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=True, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
            #print(list_sec3_org_table_CO[i])
            table_draw = ARC.report2_table_nons(list_sec3_org_table_CO[i])
            table_draw.wrapOn(c, 500, 300)
            #table_draw.drawOn(c, 4.0*inch, ((list_tbl_y[0]+3.5)-((len(list_sec3_org_table_CO[i])*0.25)+0.5))*inch)
            table_draw.drawOn(c, 4.0*inch, ((list_tbl_y[0])-((len(list_sec3_org_table_CO[i])*0.25)+0.5))*inch)
            if bIsOrgfitonepage == True:
                s = ARC.get_atbnote(list_sec3_atbnote_CO[i])
                snote = [section3_note_1 + section3_note_2_1 + ("; " if len(s)>0 else " ") + s]
                ARC.report_context(c,snote, 1.0*inch,0.4*inch,460,120, font_size=9,line_space=12)
                
                canvas_printpage(c, ipage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
                ipage = ipage+1
            #HO
            y = 1 #specify which print position in y position list to be used
            if bIsOrgfitonepage == True:
                y=0
            ARC.report_title(c,'Section [3]: AMR proportion report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
            ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
            line_1_1 = ["<b>" + "Blood: " + lst_org_format[i] + "</b>"]
            line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
            line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(lst_numpat_HO[i]) + " )" + bold_blue_ital_ed]
            ARC.report_context(c,line_1_1, 1.0*inch, list_head_y[y]*inch, 300, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_2, 4.0*inch, list_head_y[y]*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_3, 5.5*inch, list_head_y[y]*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            c.drawImage(path_result + 'Report3_AMR_' + lst_org_short[i] + "_Hospital.png", 0.3*inch, (list_fig_y[y]+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=True, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
            table_draw = ARC.report2_table_nons(list_sec3_org_table_HO[i])
            table_draw.wrapOn(c, 500, 300)
            #table_draw.drawOn(c, 4.0*inch, ((list_tbl_y[y]+3.5)-((len(list_sec3_org_table_HO[i])*0.25)+0.5))*inch)
            table_draw.drawOn(c, 4.0*inch, ((list_tbl_y[y])-((len(list_sec3_org_table_HO[i])*0.25)+0.5))*inch)
            s = ARC.get_atbnote(list_sec3_atbnote_HO[i])
            snote = [section3_note_1 + section3_note_2_1 +  ("; " if len(s)>0 else " ") + s]
            ARC.report_context(c,snote, 1.0*inch,0.4*inch,460,120, font_size=9,line_space=12)
            canvas_printpage(c, ipage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
            ipage = ipage+1
        except Exception as e:
             AL.printlog("Error SECTION 3 at " + lst_org_format[i] +  " : " + str(e),True,logger) 
             logger.exception(e)
             pass
    
def section4(c,logger,result_table, result_blo_table, result_pat_table,sec4_pat,
             startpage, lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variable
    iden1_op = "<para leftindent=\"35\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    blo_num = result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0]
    pat_num_pos_blo = result_table.loc[result_table["Parameters"]=="Number_of_patients_sampled_for_blood_culture","Values"].tolist()[0]
    ##Page1
    section4_page1_1_1 = "For each pathogen and antibiotic under surveillance, the frequencies of patients with new infections are calculated per 100,000 tested patients."
    section4_page1_1 = [section4_page1_1_1]
    section4_page1_2_1 = "The microbiology data file had:" #3.1 3104
    section4_page1_2_2 = "<i>" + "Specimen collection dates ranged from " +"</i>" +\
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        "<i>" + " to " + "</i>" + \
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section4_page1_2_3 = "Number of records on blood specimens collected within the above date range:" #3.1 3104
    section4_page1_2_4 = blo_num +" blood specimen records"
    section4_page1_2_5 = "*Number of patients sampled for blood culture within the above date range:"
    section4_page1_2_6 = pat_num_pos_blo + " patients sampled for blood culture"
    section4_page1_2 = [section4_page1_2_1, 
                        iden1_op + section4_page1_2_2 + iden_ed, 
                        iden1_op + "<i>" + section4_page1_2_3 +"</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op +  section4_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden1_op + "<i>" + section4_page1_2_5 +"</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section4_page1_2_6 + bold_blue_ital_ed + iden_ed]
    section4_page1_3_1 = "*Number of patients sampled for blood culture is used as denominator to estimate the frequency of infections per 100,000 tested patients"
    section4_page1_3 = [green_op + section4_page1_3_1 + green_ed]
    section4_page1_4_1 = "The following figures show the frequncy of infections for patients with blood culture tested."
    section4_page1_4 = [section4_page1_4_1]
    ##Page2-3
    section4_page2_1 = "*Frequency of infection per 100,000 tested patients represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested patients (denominator). " + \
                        "AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
                        
    section4_page2_2 = "CI=confidence interval; NS=non−susceptible; NA=not available/reported/tested"
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section4_page2_2 = "CI=confidence interval; R=resistant; NA=not available/reported/tested"
    
    s = ARC.get_atbnoteper_priority_pathogen(sec4_pat)
    section4_page2 = [section4_page2_1, section4_page2_2]
    section4_page2_2 = section4_page2_2 + " " + ("; " if len(s) > 0 else " ") + s
    section4_page3 = [section4_page2_1, section4_page2_2]
    ######### SECTION4: PAGE1 #########
    ARC.report_title(c,'Section [4]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section4_page1_1, 1.0*inch, 6.6*inch, 460, 200, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 7.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section4_page1_2, 1.0*inch, 5.25*inch, 460, 150, font_size=11)
    ARC.report_title(c,'Note',1.07*inch, 4.5*inch,'darkgreen',font_size=12)
    ARC.report_context(c,section4_page1_3, 1.0*inch, 3.65*inch, 460, 50, font_size=11)
    ARC.report_context(c,section4_page1_4, 1.0*inch, 2.8*inch, 460, 50, font_size=11)
    canvas_printpage(c, startpage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION4: PAGE2 #########
    ARC.report_title(c,'Section [4]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_pos_blo) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.3*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.3*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_blo_table))
    c.drawImage(path_result+"Report4_frequency_blood.png", 0.7*inch, (6.0+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_blo_table)
    table_draw.wrapOn(c, 230, 300)
    #table_draw.drawOn(c, 4.3*inch, 5.5*inch)
    table_draw.drawOn(c, 4.3*inch, ((6.42+3.5)-((len(result_blo_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section4_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+1, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION4: PAGE3 #########
    ARC.report_title(c,'Section [4]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        line_1_1 = ["<b>" + "Blood: " + " Resistant pathogens" + "</b>"]    
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_pos_blo) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.3*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.3*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_pat_table))
    c.drawImage(path_result+"Report4_frequency_pathogen.png", 0.7*inch, (6.0+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_pat_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, ((6.42+3.5)-((len(result_pat_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section4_page3, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+2, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 

def section5(c,logger,result_table, result_com_table, result_hos_table, result_com_amr_table, result_hos_amr_table,sec5_pat,
             startpage, lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start     = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end     = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    blo_num     = result_table.loc[result_table["Parameters"]=="Number_of_blood_specimens_collected","Values"].tolist()[0]
    pat_num_pos_blo     = result_table.loc[result_table["Parameters"]=="Number_of_patients_sampled_for_blood_culture","Values"].tolist()[0]
    pat_num_w2day   = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_blood_culture_within_first_2_days_of_admission","Values"].tolist()[0]
    pat_num_wo2day  = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_blood_culture_within_after_2_days_of_admission","Values"].tolist()[0]
    pat_num_unk     = result_table.loc[result_table["Parameters"]=="Number_of_patients_with_unknown_origin","Values"].tolist()[0]
    pat_num_oth     = result_table.loc[result_table["Parameters"]=="Number_of_patients_had_more_than_one_admission","Values"].tolist()[0]
    ##Page1
    section5_page1_1_1 = "For each infection origin, pathogen and antibiotic under surveillance, the frequencies of patients with new infections are calculated per 100,000 tested patients."
    section5_page1_1 = [section5_page1_1_1]
    section5_page1_2_1 = "The data included in the analysis had:"
    section5_page1_2_2 = "<i>" + "Specimen collection dates ranged from " + "</i>" + \
                        bold_blue_ital_op + spc_date_start + bold_blue_ital_ed + \
                        " to " + \
                        bold_blue_ital_op + spc_date_end + bold_blue_ital_ed
    section5_page1_2_3 = "Number of records on blood specimens collected within the above date range:" #3.1 3104
    section5_page1_2_4 = blo_num + " blood specimen records"
    section5_page1_2_5 = "Number of patients sampled for blood culture within the above date range:"
    section5_page1_2_6 = pat_num_pos_blo + " patients sampled for blood culture"
    section5_page1_2_7 = bold_blue_ital_op + pat_num_w2day + bold_blue_ital_ed + \
                        "<i>" + " patients had at least one admission having the first blood culture drawn within first 2 calendar days of hospital admission." + "</i>"
    section5_page1_2_8 = "This parameter is used as a denominators for frequency of community−origin bacteraemia (per 100,000 patients tested for blood culture on admission)."
    section5_page1_2_9 = bold_blue_ital_op + pat_num_wo2day + bold_blue_ital_ed + \
                        "<i>" + " patients had at least one admission having the first blood culture drawn after 2 calendar days of hospital admission." + "</i>"
    section5_page1_2_10 = "This parameter is used as a denominators for frequency of hospital−origin bacteraemia (per 100,000 patients tested for blood culture for HAI)."
    section5_page1_2_11 = bold_blue_ital_op + pat_num_unk + bold_blue_ital_ed + \
                        "<i>" + " patients had a blood drawn for culture and with unknown origin of infection." + "</i>"
    section5_page1_2_12 = "Validation of this statistics is highly recommended."
    section5_page1_2 = [section5_page1_2_1, 
                        iden1_op + section5_page1_2_2 + iden_ed, 
                        iden1_op + "<i>" + section5_page1_2_3+ "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section5_page1_2_4 + bold_blue_ital_ed + iden_ed, 
                        iden1_op + "<i>" + section5_page1_2_5+ "</i>" + iden_ed, 
                        iden1_op + bold_blue_ital_op + section5_page1_2_6 + bold_blue_ital_ed + iden_ed, 
                        iden2_op + add_blankline + section5_page1_2_7 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_8 + "</i>" + iden_ed, 
                        iden2_op + section5_page1_2_9 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_10 + "</i>" + iden_ed, 
                        iden2_op + section5_page1_2_11 + iden_ed, 
                        iden2_op + "<i>" + section5_page1_2_12 + "</i>" + iden_ed]
    section5_page1_3_1 = bold_blue_ital_op + pat_num_oth + bold_blue_ital_ed + \
                        "<i>" + green_op + " patients had more than one admissions, of which at least one admission had the first blood culture drawn within the first 2 calendar days of hospital admission AND at least one admission had the first blood culture drawn after 2 calendar days of hospital admission." + \
                        green_ed + "</i>"
    section5_page1_3 = [iden2_op + section5_page1_3_1 + iden_ed]
    section5_page1_4_1 = "The following figures show the frequency of infections for patients with blood culture tested and stratified by infection origin, under this surveillance."
    section5_page1_4 = [section5_page1_4_1]
    ##Page2-5
    section5_page2_1 = "*Frequency of infection per 100,000 tested patients on admission represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested population on admission (denominator). " + \
                        "AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_1v2= "*Frequency of infection per 100,000 tested population at risk of HAI represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested population at risk of HAI (denominator). AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_1v3= "*Frequency of infection per 100,000 tested patients represents the number of patients with blood culture positive for a pathogen (numerator) over the total number of tested patients (denominator). AMASS application de−duplicates the data by included only the first isolate of each patient per specimen type per reporting period."
    section5_page2_2 = "CI=confidence interval; NA=not available/reported/tested"
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section5_page2_2 = "CI=confidence interval; NA=not available/reported/tested"
    section5_page2 = [section5_page2_1 + " " + section5_page2_2]
    s = ARC.get_atbnoteper_priority_pathogen(sec5_pat)
    section5_page2v2 = [section5_page2_1v2 + " " + section5_page2_2]
    section5_page2v3 = [section5_page2_1v3 + " " + section5_page2_2 + ("; " if len(s) > 0 else " ") + s]
    section5_page2v4 = [section5_page2_1 + " " +  section5_page2_2+ ("; " if len(s) > 0 else " ") + s]
    ######### SECTION5: PAGE1 #########
    ARC.report_title(c,'Section [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section5_page1_1, 1.0*inch, 8.0*inch, 460, 100, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 7.9*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section5_page1_2, 1.0*inch, 3.2*inch, 480, 330, font_size=11)
    ARC.report_title(c,'Note:',2.0*inch, 3.0*inch,'darkgreen',font_size=12)
    ARC.report_context(c,section5_page1_3, 1.0*inch, 1.5*inch, 460, 100, font_size=11)
    ARC.report_context(c,section5_page1_4, 1.0*inch, 1.0*inch, 460, 50, font_size=11)
    canvas_printpage(c, startpage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION5: PAGE2 #########
    ARC.report_title(c,'Section [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Community-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_w2day) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_com_table))    
    #c.drawImage(path_result+"Report5_incidence_community.png", 0.7*inch, 2.7*inch, preserveAspectRatio=False, width=3.5*inch, height=6.5*inch,showBoundary=False) 
    c.drawImage(path_result+"Report5_incidence_community.png", 0.7*inch, (5.8+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_com_table)
    table_draw.wrapOn(c, 230, 300)
    #table_draw.drawOn(c, 4.3*inch, 5.2*inch)
    table_draw.drawOn(c, 4.3*inch, ((6.3+3.5)-((len(result_com_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section5_page2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+1, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION5: PAGE3 #########
    ARC.report_title(c,'Section [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_wo2day) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_hos_table))
    c.drawImage(path_result+"Report5_incidence_hospital.png", 0.7*inch, (5.8+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_hos_table)
    table_draw.wrapOn(c, 230, 300)
    table_draw.drawOn(c, 4.3*inch, ((6.3+3.5)-((len(result_hos_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section5_page2v2, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+2, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION5: PAGE4 #########
    ARC.report_title(c,'Section [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        line_1_1 = ["<b>" + "Blood: " + " Resistant pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Community-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_w2day) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_com_amr_table))
    c.drawImage(path_result+"Report5_incidence_community_antibiotic.png", 0.7*inch, (5.8+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_com_amr_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, ((6.3+3.5)-((len(result_com_amr_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section5_page2v4, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+3, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION5: PAGE5 #########
    ARC.report_title(c,'Section [5]: AMR frequency report',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'with stratification by infection origin',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    line_1_1 = ["<b>" + "Blood: " + " Non-susceptible pathogens" + "</b>"]
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        line_1_1 = ["<b>" + "Blood: " + " Resistant pathogens" + "</b>"]
    line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
    line_1_3 = [bold_blue_ital_op + " ( No. of patients = " + str(pat_num_wo2day) + " )" + bold_blue_ital_ed]
    ARC.report_context(c,line_1_1, 1.0*inch, 9.0*inch, 300, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_2, 4.0*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ARC.report_context(c,line_1_3, 5.5*inch, 9.0*inch, 200, 50, font_size=12, font_align=TA_LEFT)
    ifig_dis_H = ARC.cal_sec4_fig_height(len(result_hos_amr_table))
    c.drawImage(path_result+"Report5_incidence_hospital_antibiotic.png", 0.7*inch, (5.8+(3.5 - ifig_dis_H))*inch, preserveAspectRatio=False, width=3.5*inch, height=ifig_dis_H*inch,showBoundary=False) 
    table_draw = ARC.report2_table_nons(result_hos_amr_table)
    table_draw.wrapOn(c, 240, 300)
    table_draw.drawOn(c, 4.3*inch, ((6.3+3.5)-((len(result_hos_amr_table)*0.45)+0.5))*inch)
    ARC.report_context(c,section5_page2v3, 1.0*inch, 0.4*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c, startpage+4, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    
def section6(c,logger,result_table, sec6_mor, result_mor_table,list_sec6_mor_tbl_CO,list_sec6_mor_tbl_HO,
             lst_org, lst_org_short, lst_org_full, df_numpat_com, df_numpat_hos,
             startpage, lastpage="47", today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"

    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"

    add_blankline = "<br/>"
    ##variables
    spc_date_start = result_table.loc[(result_table["Type_of_data_file"]=="microbiology_data")&(result_table["Parameters"]=="Minimum_date"),"Values"].tolist()[0]
    spc_date_end       = result_table.loc[(result_table["Type_of_data_file"]=="microbiology_data")&(result_table["Parameters"]=="Maximum_date"),"Values"].tolist()[0]
    pat_num_pos_org    = result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive_for_organism_under_this_survey","Values"].tolist()[0]
    pat_num_pos_org_com= result_table.loc[result_table["Parameters"]=="Number_of_patients_with_community_origin_BSI","Values"].tolist()[0]
    pat_num_pos_org_hos= result_table.loc[result_table["Parameters"]=="Number_of_patients_with_hospital_origin_BSI","Values"].tolist()[0]
    hos_date_start     = result_table.loc[(result_table["Type_of_data_file"]=="hospital_admission_data")&(result_table["Parameters"]=="Minimum_date"),"Values"].tolist()[0]
    hos_date_end       = result_table.loc[(result_table["Type_of_data_file"]=="hospital_admission_data")&(result_table["Parameters"]=="Maximum_date"),"Values"].tolist()[0]
    hos_num            = result_table.loc[result_table["Parameters"]=="Number_of_records","Values"].tolist()[0]
    pat_num_hos        = result_table.loc[result_table["Parameters"]=="Number_of_patients_included","Values"].tolist()[0]
    pat_num_dead       = result_table.loc[result_table["Parameters"]=="Number_of_deaths","Values"].tolist()[0]
    per_mortal         = result_table.loc[result_table["Parameters"]=="Mortality","Values"].tolist()[0]
    ##Page1
    section6_page1_1_1 = "A surveillance report on mortality involving AMR infections and antimicrobial−susceptible infections with stratification by origin of infection is generated only if data on patient outcomes (i.e. discharge status) are available. " + \
                    "Antimicrobial−resistant infection is a threat to modern health care, and the impact of the infection on patient outcomes is largely unknown. " + \
                    "Performing analyses and generating reports on mortality often takes time and resources."
    section6_page1_1_2 = "The term \"mortality involving AMR and antimicrobial−susceptible infections\" was used because the mortality reported was all−cause mortality. " + \
                        "This measure of mortality included deaths caused by or related to other underlying and intermediate causes." #3.1 3104
    section6_page1_1_3 = "Here, AMASS summarized the overall mortality of patients with antimicrobial−resistant and antimicrobial−susceptible bacteria bloodstream infections (BSI)."
    section6_page1_1 = [section6_page1_1_1, 
                        add_blankline + section6_page1_1_2, 
                        add_blankline + section6_page1_1_3]
    section6_page1_2_1 = "The data included in the analysis had:"
    section6_page1_2_2 = "Sample collection dates ranged from "+ \
                        bold_blue_ital_op + str(spc_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(spc_date_end) + bold_blue_ital_ed
    section6_page1_2_3 = "Number of patients with blood culture positive for the origanism under the surveillance:"
    section6_page1_2_4 = str(pat_num_pos_org) + " patients"
    section6_page1_2_5 = "Number of patients with community−origin BSI:"
    section6_page1_2_6 = str(pat_num_pos_org_com) + " patients"
    section6_page1_2_7 = "Number of patients with hospital−origin BSI:"
    section6_page1_2_8 = str(pat_num_pos_org_hos) + " patients"
    section6_page1_2_9 = "The hospital admission data file had:"
    section6_page1_2_10 = "Hospital admission dates ranging from "+ \
                        bold_blue_ital_op + str(hos_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(hos_date_end) + bold_blue_ital_ed
    section6_page1_2_11 = "Number of records in the raw hospital admission data:"
    section6_page1_2_12 = str(hos_num) + " records"
    section6_page1_2_13 = "Number of patients included in the analysis (de−duplicated):"
    section6_page1_2_14 = str(pat_num_hos) + " patients"
    section6_page1_2_15 = "Number of patients having death as an outcome in any admission data records:"
    section6_page1_2_16 = str(pat_num_dead) + " patients"
    section6_page1_2_17 = "Overall mortality:"
    try:
        section6_page1_2_18 = str(int(per_mortal))
    except:
        section6_page1_2_18 = str(per_mortal)
        pass 
    
    section6_page1_2 = [section6_page1_2_1, 
                        iden1_op + "<i>" + section6_page1_2_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_3 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_4 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_5 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_6 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_7 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_8 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        section6_page1_2_9, 
                        iden1_op + "<i>" + section6_page1_2_10 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_11 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_12 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden1_op + "<i>" + section6_page1_2_13 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + bold_blue_ital_op + section6_page1_2_14 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_15 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_16 + bold_blue_ital_ed + "</i>" + iden_ed, 
                        iden2_op + "<i>" + section6_page1_2_17 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + bold_blue_ital_op + section6_page1_2_18 + bold_blue_ital_ed + "</i>" + iden_ed]
    
    section6_page2_org = ["<b>" + "Organism" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>", 
                        "<b>" + "" + "</b>"]
    for i in range(len(lst_org)):
        section6_page2_org.append("<b>" + lst_org[i] + "</b>")
    section6_page2_org.append("<b>" + "Total:" + "</b>")
    ##Page3-8
    section6_page2_1_1 = "AMASS application merged the microbiology data file and hospital admission data file. " + \
                        "The merged dataset was then de−duplicated so that only the first isolate per patient per specimen per reporting period was included in the analysis. " + \
                        "The de−duplicated data was stratified by infection origin (community−origin infection or hospital−origin infection)."
    section6_page2_1 = [section6_page2_1_1]
    section6_page2_2_1 = "The following figures and tables show the mortality of patients who were blood culture positive for antimicrobial non−susceptible and susceptible isolates."
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        section6_page2_2_1 = "The following figures and tables show the mortality of patients who were blood culture positive for antimicrobial resistant and susceptible isolates."
    section6_page2_2 = [section6_page2_2_1]
    ##Page3
    stemp = "NS=non−susceptible; S=susceptible; CI=confidence interval"
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        stemp = "R=resistant; S=susceptible (including sensitive and intermediate categories); CI=confidence interval" 
    section6_page3_1_1 = "*Mortality is the proportion (%) of in−hospital deaths (all−cause deaths). " + \
                        "This represents the number of in−hospital deaths (numerator) over the total number of patients with blood culture positive for the organism and the type of pathogen (denominator). " + \
                        "AMASS application de−duplicates the data by included only the first isolate per patient per specimen type per evaluation period. " + \
                        stemp
    ######### SECTION6: PAGE1 #########
    ARC.report_title(c,'Section [6] Mortality involving AMR and',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.6*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section6_page1_1, 1.0*inch, 6.1*inch, 460, 250,font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 6.0*inch,'#3e4444',font_size=12)
    ARC.report_context(c,section6_page1_2, 1.0*inch, 0.67*inch, 460, 380,font_size=11)
    canvas_printpage(c, startpage, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION6: PAGE2 #########
    ARC.report_context(c,section6_page2_1, 1.0*inch, 8.5*inch, 460, 120,font_size=11)
    #ARC.report_context(c,section6_page2_org, 1.0*inch, 5.0*inch, 460, 250, font_size=11, line_space=18.0)
    table_draw = ARC.report6_table(result_mor_table)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 5.9*inch)
    ARC.report_context(c,section6_page2_2, 1.0*inch, 4.0*inch, 460, 50,font_size=11)
    canvas_printpage(c, startpage+1, lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### SECTION3: PAGE3-12 #########   
    iPage = startpage+1
    bIsStartofnewpage = True
    bPrintPage = False
    iPos_head_y_CO = 0
    iPos_head_y_HO = 0
    iPos_fig_y_CO = 0
    iPos_fig_y_HO = 0
    iPos_tbl_y_CO = 0
    iPos_tbl_y_HO = 0
    #scurenote = ""
    list_curnote = []
    for i in range(len(lst_org)):
        try:
            tmp_df = sec6_mor[sec6_mor["Organism"].str.lower()==str(lst_org_full[i]).lower()]
            ifig_dis_H_CO = ARC.cal_sec6_fig_height(len(list_sec6_mor_tbl_CO[i]))
            ifig_dis_H_HO = ARC.cal_sec6_fig_height(len(list_sec6_mor_tbl_HO[i]))
            #ifig_dis_offset = 3.5 - ifig_dis_H
            #Print position calculation
            bPrintPage = False
            if (len(list_sec6_mor_tbl_CO[i]) + len(list_sec6_mor_tbl_HO[i])) > AC.CONST_MAX_ATBCOUNTFITHALFPAGE_MORALITY:
                #Fullpage both CO and HO should be bigger
                iPos_head_y_CO = 9.0
                iPos_head_y_HO = 5.4
                iPos_fig_y_CO = 6.1+3.5 - ifig_dis_H_CO
                iPos_fig_y_HO = 2.5+3.5 - ifig_dis_H_HO
                iPos_tbl_y_CO = (6.1+3.6)-((len(list_sec6_mor_tbl_CO[i])*0.25)+0.5)
                iPos_tbl_y_HO = (2.5+3.6)-((len(list_sec6_mor_tbl_HO[i])*0.25)+0.5)
                #scurenote = ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                list_curnote = ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                bPrintPage = True
                bIsStartofnewpage = True
            elif bIsStartofnewpage: 
                #Half top page
                iPos_head_y_CO = 9.0
                iPos_head_y_HO = 7.2
                iPos_fig_y_CO = 6.1+3.5 - ifig_dis_H_CO
                iPos_fig_y_HO = 4.3+3.5 - ifig_dis_H_HO
                iPos_tbl_y_CO = (6.1+3.6)-((len(list_sec6_mor_tbl_CO[i])*0.25)+0.5)
                iPos_tbl_y_HO = (4.3+3.6)-((len(list_sec6_mor_tbl_HO[i])*0.25)+0.5)
                #scurenote = ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                list_curnote = ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                #Check if next require fullpage then print page (note + page no)
                ii = i+1
                if ii>= len(lst_org):
                    bPrintPage = True
                else:
                    if (len(list_sec6_mor_tbl_CO[ii]) + len(list_sec6_mor_tbl_HO[ii])) <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE_MORALITY:
                        bPrintPage = False
                    else:
                        bPrintPage = True
                bIsStartofnewpage = False
            else:
                #Half bottom page  
                iPos_head_y_CO = 5.4
                iPos_head_y_HO = 3.6
                iPos_fig_y_CO = 2.5+3.5 - ifig_dis_H_CO
                iPos_fig_y_HO = 0.7+3.5 - ifig_dis_H_HO
                iPos_tbl_y_CO = (2.5+3.6)-((len(list_sec6_mor_tbl_CO[i])*0.25)+0.5)
                iPos_tbl_y_HO = (0.7+3.6)-((len(list_sec6_mor_tbl_HO[i])*0.25)+0.5)
                #list_curnote = list_curnote+list_sec2_atbnote[i]
                #list_curnote = list(set(list_curnote))
                #scurenote = scurenote + ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                list_curnote = list_curnote + ARC.get_atbnoteper_priority_pathogen_sec6(tmp_df)
                bPrintPage = True
                bIsStartofnewpage = True
            #Print content
            ARC.report_title(c,'Section [6] Mortality involving AMR and',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
            ARC.report_title(c,'antimicrobial−susceptible infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
            #Print CO
            line_1_1 = ["<b>" + "Blood: " + lst_org[i] + "</b>"]
            line_1_2 = ["<b>" + "Community-origin" + "</b>"]
            line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(df_numpat_com.loc[lst_org_full[i]]) + " )" + bold_blue_ital_ed]
            ARC.report_context(c,line_1_1, 1.0*inch, iPos_head_y_CO*inch, 300, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_2, 4.0*inch, iPos_head_y_CO*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_3, 5.5*inch, iPos_head_y_CO*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            c.drawImage(path_result+'Report6_mortality_'+ lst_org_short[i] + "_" + "community"+".png", 1.0*inch, iPos_fig_y_CO*inch, preserveAspectRatio=True, width=3.5*inch, height=ifig_dis_H_CO*inch,showBoundary=False) 
            table_draw = ARC.report2_table_nons(list_sec6_mor_tbl_CO[i])
            table_draw.wrapOn(c, 500, 300)
            table_draw.drawOn(c, 4.7*inch, iPos_tbl_y_CO*inch)
            #Print HO
            line_1_1 = ["<b>" + "Blood: " + lst_org[i] + "</b>"]
            line_1_2 = ["<b>" + "Hospital-origin" + "</b>"]
            line_1_3 = [bold_blue_ital_op + "( No. of patients = " + str(df_numpat_hos.loc[lst_org_full[i]]) + " )" + bold_blue_ital_ed]
            ARC.report_context(c,line_1_1, 1.0*inch, iPos_head_y_HO*inch, 300, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_2, 4.0*inch, iPos_head_y_HO*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            ARC.report_context(c,line_1_3, 5.5*inch, iPos_head_y_HO*inch, 200, 50, font_size=12, font_align=TA_LEFT)
            c.drawImage(path_result+'Report6_mortality_'+ lst_org_short[i] + "_" + "hospital"+".png", 1.0*inch, iPos_fig_y_HO*inch, preserveAspectRatio=True, width=3.5*inch, height=ifig_dis_H_HO*inch,showBoundary=False) 
            table_draw = ARC.report2_table_nons(list_sec6_mor_tbl_HO[i])
            table_draw.wrapOn(c, 500, 300)
            table_draw.drawOn(c, 4.7*inch, iPos_tbl_y_HO*inch)
            #Print page (note + page no) if needed
            if bPrintPage:
                #scurenote  =  section6_page3_1_1 + "; " + scurenote 
                s = ARC.get_atbnote_sec6(list_curnote)
                section6_page3 = section6_page3_1_1 + ("; " if len(s) > 0 else " ") + s
                #ARC.report_context(c,[scurenote] ,1.0*inch,0.4*inch,460,120, font_size=9,line_space=12)
                ARC.report_context(c,[section6_page3] ,1.0*inch,0.4*inch,460,120, font_size=9,line_space=12)
                iPage = iPage + 1
                canvas_printpage(c,iPage,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
                #scurenote  = ""
                list_curnote = []
                bIsStartofnewpage = True
        except Exception as e:
             AL.printlog("Error SECTION 6 at " + lst_org_full[i] +  " : " + str(e),True,logger) 
             logger.exception(e)
             pass
    
def annexA_A1(c,logger,result_table, org_table, pat_table,result_table_A11,  pat_table_A11,
                    startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    iden1_op = "<para leftindent=\"35\">"
    iden2_op = "<para leftindent=\"70\">"
    iden_ed = "</para>"
    bold_blue_ital_op = "<b><i><font color=\"#000080\">"
    bold_blue_ital_ed = "</font></i></b>"
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    spc_date_start  = result_table.loc[result_table["Parameters"]=="Minimum_date","Values"].tolist()[0]
    spc_date_end    = result_table.loc[result_table["Parameters"]=="Maximum_date","Values"].tolist()[0]
    spc_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_all_culture_positive","Values"].tolist()[0]
    blo_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_blood_culture_positive","Values"].tolist()[0]
    csf_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_csf_culture_positive","Values"].tolist()[0]
    gen_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_genital_swab_culture_positive","Values"].tolist()[0]
    res_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_rts_culture_positive","Values"].tolist()[0]
    sto_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_stool_culture_positive","Values"].tolist()[0]
    uri_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_urine_culture_positive","Values"].tolist()[0]
    oth_num_pos_plus= result_table.loc[result_table["Parameters"]=="Number_of_others_culture_positive","Values"].tolist()[0]
    
    ##Page1
    #annexA_page1_1_1 = "This supplementary report has two parts; including (A1) notifiable bacterial infections and (A2) mortality involving notifiable bacterial infections. The AMR proportion notifiable bacterial infections supplementary report is generated by default, even if the hospital admission data file is unavailable. This is to enable hospitals with only microbiology data available to utilize the de-duplication and report generation functions of AMASS."
    annexA_page1_1_1 = "This supplementary report has two parts; including (A1) notifiable bacterial infections and (A2) mortality involving notifiable bacterial infections."
    annexA_page1_1_2 = "Please note that the completion of this supplementary report is strongly dependent on the availability of data (particularly, data of all bacterial pathogens and all specimen types) and the completion of the data dictionary files to make sure that AMASS can understand each notifiable bacterium and each type of specimen."
    #annexA_page1_1_3 = "Annex A includes various type of specimens including blood, cerebrospinal fluid (CSF), respiratory tract specimens, urine, genital swab, stool and other or unknown sample types. The microorganisms in this report were initially selected from common notifiable bacterial diseases in Thailand."
    annexA_page1_1_3 = "Annex A includes all specimens types including blood, cerebrospinal fluid (CSF), respiratory tract specimens, urine, genital swab, stool and other or unknown sample types. The notifiable bacteria included in this report were initially selected from common notifiable bacterial diseases in Thailand."
    annexA_page1_1_4 = "Annex A1a is generated by default, even if the hospital admission data file is unavailable. It may include patients who were not admitted to the evaluation hospital during the evaluation period. Annex A1b and Annex A2 are generated only if hospital admission data file is available, and include only patients who had a clinical specimen culture positive for a notifiable bacterium during their admissions to the evaluation hospital during the evaluation period."
    annexA_page1_1 = [annexA_page1_1_1, 
                    add_blankline + annexA_page1_1_2, 
                    add_blankline + annexA_page1_1_3,
                    add_blankline + annexA_page1_1_4]

    annexA_page1_2_1 = "Note: The list of notifiable bacteria included in AMASS was generated based on the literature review and the collaboration with Department of Disease Control, Ministry of Public Health, Thailand. The list could be expanded or modified in future versions of AMASS.."
    annexA_page1_2 = [green_op + annexA_page1_2_1 + green_ed]
    ##Page2
    annexA_page2_1_1 = "The microbiology data file had:" #3.1 3104
    annexA_page2_1_2 = "Specimen collection dates ranged from " + \
                        bold_blue_ital_op + str(spc_date_start) + bold_blue_ital_ed + \
                        "  to  " + \
                        bold_blue_ital_op + str(spc_date_end) + bold_blue_ital_ed
    annexA_page2_1_3 = "Number of records of specimens culture positive for a notifiable bacterium under the surveillance:"
    annexA_page2_1_4 = bold_blue_ital_op + str(spc_num_pos_plus) + bold_blue_ital_ed + "  specimen records (" + \
                        bold_blue_ital_op + str(blo_num_pos_plus) + " , " + str(csf_num_pos_plus) + " , " + \
                        str(gen_num_pos_plus) + " , " + str(res_num_pos_plus) + " , " + str(sto_num_pos_plus) + " , " + \
                        str(uri_num_pos_plus) + " , " + str(oth_num_pos_plus) + bold_blue_ital_ed + \
                        " were blood, CSF, genital swab, respiratory tract specimens, stool, urine, and other or unknown sample types, respectively) "
    annexA_page2_1_5 = "AMASS de-duplicated the data by including only the first isolate per patient per specimen type per evaluation period as described in the method. The number of patients with notifiable bacterial infections is as follows:"
    annexA_page2_1 = [annexA_page2_1_1, 
                    iden1_op + "<i>" + annexA_page2_1_2 + "</i>" + iden_ed, 
                    iden1_op + "<i>" + annexA_page2_1_3 + "</i>" + iden_ed, 
                    iden2_op + "<i>" + annexA_page2_1_4 + "</i>" + iden_ed, 
                    add_blankline + annexA_page2_1_5]
    annexA_page2_2_1 = "*Some patients may have more than one type of specimen culture positive for a notifiable bacterium. Some patients may have clinical specimens culture positive for multiple notifiable bacteria. Some patients may not be hospitalized at the survey hospital during the evaluation period."
    annexA_page2_2_2 = "CSF = Cerebrospinal fluid; RTS = Respiratory tract specimens; Others = Other or unknown sample types; NA = Not applicable (i.e. the specimen type is not available or identified in the microbiology data file)"
    #annexA_page2_2_3 = "NA = Not applicable (i.e. the specimen type is not available or identified in the microbiology data file)"
    #annexA_page2_2 = [annexA_page2_2_1,
    #                annexA_page2_2_2, 
    #                annexA_page2_2_3]
    annexA_page2_2 = [annexA_page2_2_1,
                    annexA_page2_2_2]
    ######### ANNEX A: PAGE1 ##########
    ARC.report_title(c,'Annex A: Supplementary report on notifiable bacterial',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'infections',1.07*inch, 10.2*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,annexA_page1_1, 1.0*inch, 4.5*inch, 460, 350,font_size=11)
    ARC.report_title(c,'Notifiable bacteria under the surveillance',1.07*inch, 4.8*inch,'#3e4444',font_size=12)
    table_draw = ARC.report_table_annexA_page1(org_table)
    table_draw.wrapOn(c, 700, 700)
    table_draw.drawOn(c, 1.2*inch, 3.0*inch)
    ARC.report_context(c,annexA_page1_2, 1.0*inch, 0.5*inch, 460, 100,font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### ANNEX A: PAGE2 ##########
    ARC.report_title(c,'Annex A1a: Notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Results',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,annexA_page2_1, 1.0*inch, 5.9*inch, 460, 250,font_size=11)
    table_draw = ARC.report_table_annexA_page2(pat_table)
    table_draw.wrapOn(c, 400, 300)
    if len(pat_table) < 7:
        table_draw.drawOn(c, 1.2*inch, 3.2*inch)
    else:
        table_draw.drawOn(c, 1.2*inch, 2.7*inch)
    ARC.report_context(c,annexA_page2_2, 1.0*inch, 0.7*inch, 460, 100,font_size=9,line_space=12)
    canvas_printpage(c,startpage+1,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ######### ANNEX A1.1: PAGE2A ##########
    bgotdata_annexA11 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secA_res_i_A11)
    if bgotdata_annexA11:
        hos_date_start_A11  = result_table_A11.loc[result_table_A11["Parameters"]=="Minimum_admdate","Values"].tolist()[0]
        hos_date_end_A11    = result_table_A11.loc[result_table_A11["Parameters"]=="Maximum_admdate","Values"].tolist()[0]
        spc_date_start_A11  = result_table_A11.loc[result_table_A11["Parameters"]=="Minimum_date","Values"].tolist()[0]
        spc_date_end_A11    = result_table_A11.loc[result_table_A11["Parameters"]=="Maximum_date","Values"].tolist()[0]
        spc_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_all_culture_positive","Values"].tolist()[0]
        blo_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_blood_culture_positive","Values"].tolist()[0]
        csf_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_csf_culture_positive","Values"].tolist()[0]
        gen_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_genital_swab_culture_positive","Values"].tolist()[0]
        res_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_rts_culture_positive","Values"].tolist()[0]
        sto_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_stool_culture_positive","Values"].tolist()[0]
        uri_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_urine_culture_positive","Values"].tolist()[0]
        oth_num_pos_plus_A11= result_table_A11.loc[result_table_A11["Parameters"]=="Number_of_others_culture_positive","Values"].tolist()[0]
        ##Page2A (A1.1)
        annexA_page2_A11_1_1 = "After merging microbiology and hospital admission data files, and including only patients who were culture positive for a notifiable bacterium and admitted to the evaluation hospital during the evaluation period, the data used to generate the report had:"
        # annexA_page2_A11_1_2_HOS = "Hospital admission dates ranged from " + \
        #                     bold_blue_ital_op + str(hos_date_start_A11) + bold_blue_ital_ed + \
        #                     "  to  " + \
        #                     bold_blue_ital_op + str(hos_date_end_A11) + bold_blue_ital_ed
        annexA_page2_A11_1_2 = "Specimen collection dates ranged from " + \
                            bold_blue_ital_op + str(spc_date_start_A11) + bold_blue_ital_ed + \
                            "  to  " + \
                            bold_blue_ital_op + str(spc_date_end_A11) + bold_blue_ital_ed
        annexA_page2_A11_1_3 = "Number of records of specimens culture positive for a notifiable bacterium under the surveillance:"
        annexA_page2_A11_1_4 = bold_blue_ital_op + str(spc_num_pos_plus_A11) + bold_blue_ital_ed + "  specimen records (" + \
                            bold_blue_ital_op + str(blo_num_pos_plus_A11) + " , " + str(csf_num_pos_plus_A11) + " , " + \
                            str(gen_num_pos_plus_A11) + " , " + str(res_num_pos_plus_A11) + " , " + str(sto_num_pos_plus_A11) + " , " + \
                            str(uri_num_pos_plus_A11) + " , " + str(oth_num_pos_plus_A11) + bold_blue_ital_ed + \
                            " were blood, CSF, genital swab, respiratory tract specimens, stool, urine, and other or unknown sample types, respectively) "
        annexA_page2_A11_1_5 = "AMASS de-duplicated the data by including only the first isolate per patient per specimen type per evaluation period as described in the method. The number of patients with notifiable bacterial infections is as follows:"
        # annexA_page2_A11_1 = [annexA_page2_A11_1_1, 
        #                 iden1_op + "<i>" + annexA_page2_A11_1_2 + "</i>" + iden_ed, 
        #                 iden1_op + "<i>" + annexA_page2_A11_1_2_HOS + "</i>" + iden_ed, 
        #                 iden1_op + "<i>" + annexA_page2_A11_1_3 + "</i>" + iden_ed, 
        #                 iden2_op + "<i>" + annexA_page2_A11_1_4 + "</i>" + iden_ed, 
        #                 add_blankline + annexA_page2_A11_1_5]
        annexA_page2_A11_1 = [annexA_page2_A11_1_1, 
                        iden1_op + "<i>" + annexA_page2_A11_1_2 + "</i>" + iden_ed, 
                        iden1_op + "<i>" + annexA_page2_A11_1_3 + "</i>" + iden_ed, 
                        iden2_op + "<i>" + annexA_page2_A11_1_4 + "</i>" + iden_ed, 
                        add_blankline + annexA_page2_A11_1_5]
        annexA_page2_A11_2_1 = "*Some patients may have more than one type of specimen culture positive for a notifiable bacterium. Some patients may have clinical specimens culture positive for multiple notifiable bacteria."
        annexA_page2_A11_2_2 = "Some patients may have the data of a clinical specimen culture positive for a notifiable bacterium in the microbiology data file, but do not have the data of admission dates, discharge dates and/or patient outcomes in the hospital admission data file. That is the most common cause of the discrepancy between total number of patients with notifiable bacterial infections presented in the Annex A1a, A1b and A2 (followed by typos in patient identifiers in either data file)."
        annexA_page2_A11_2_3 = "CSF = Cerebrospinal fluid; RTS = Respiratory tract specimens; Others = Other or unknown sample types; NA = Not applicable (i.e. the specimen type is not available or identified in the microbiology data file)" #3.1 3104
        annexA_page2_A11_2 = [annexA_page2_A11_2_1,
                        annexA_page2_A11_2_2, 
                        annexA_page2_A11_2_3]
        ARC.report_title(c,'Annex A1b: Notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
        ARC.report_title(c,'Results',1.07*inch, 9.75*inch,'#3e4444',font_size=12)
        ARC.report_context(c,annexA_page2_A11_1, 1.0*inch, 6.15*inch, 460, 250,font_size=11)
        table_draw = ARC.report_table_annexA_page2(pat_table_A11)
        table_draw.wrapOn(c, 400, 300)
        if len(pat_table) < 7:
            table_draw.drawOn(c, 1.2*inch, 3.3*inch)
        else:
            table_draw.drawOn(c, 1.2*inch, 2.40*inch)
        ARC.report_context(c,annexA_page2_A11_2, 1.0*inch, 0.67*inch, 460, 120,font_size=9,line_space=12)
        canvas_printpage(c,startpage+2,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    else:
        annexA11_nodata_mortality(c,startpage+2,lastpage,today,ipagemode,ssectionanme,isecmaxpage)
def annexA_A2(c,logger,mor_table, startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    ##variables
    ##Page3
    annexA_page3_1_1 = "A report on mortality involving notifiable bacterial infections is generated only if data on patient outcomes (i.e. discharge status) are available. The term \"mortality involving notifiable bacterial infections\" was used because the mortality reported was all-cause mortality. This measure of mortality included deaths caused by or related to other underlying and intermediate causes. AMASS application merged the microbiology data file and hospital admission data file. The merged dataset was then de-duplicated so that only the first isolate per patient per specimen per reporting period was included in the analysis."
    annexA_page3_1 = [annexA_page3_1_1]
    annexA_page3_2_1 = "*Mortality is the proportion (%) of in-hospital deaths (all-cause deaths). This represents the number of in-hospital deaths (numerator) over the total number of patients with culture positive for each type of pathogen (denominator). Some patients may have the data of a clinical specimen culture positive for a notifiable bacterium in the microbiology data file, but do not have the data of admission dates, discharge dates and/or patient outcomes in the hospital admission data file. That is the most common cause of the discrepancy between total number of patients with notifiable bacterial infections presented in the Annex A1a, A1b and A2 (followed by typos in patient identifiers in either data file)."
    annexA_page3_2_2 = "CI = confidence interval"
    annexA_page3_2 = [annexA_page3_2_1, 
                    annexA_page3_2_2]
    ######### ANNEX A: PAGE3 ##########
    ARC.report_title(c,'Annex A2: Mortality involving notifiable bacterial infections',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,annexA_page3_1, 1.0*inch, 7.5*inch, 460, 150,font_size=11)
    ARC.report_context(c,["*Mortality (%)"], 2.0*inch, 2.2*inch, 150, 30, font_size=9, font_align=TA_CENTER, line_space=14)
    c.drawImage(path_result+"AnnexA_mortality.png", 1.2*inch, 2.5*inch, preserveAspectRatio=False, width=2.5*inch, height=5.0*inch,showBoundary=False) 
    table_draw = ARC.report_table_annexA_page3(mor_table)
    table_draw.wrapOn(c, 265, 300)
    if len(mor_table) < 7:
        table_draw.drawOn(c, 4.2*inch, 5.0*inch)
    else:
        table_draw.drawOn(c, 4.2*inch, 3.5*inch)
    ARC.report_context(c,annexA_page3_2, 1.0*inch, 0.67*inch, 460, 110,font_size=9,line_space=12)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 

def annexB(c,logger,blo_table, blo_table_bymonth, startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    add_blankline = "<br/>"
    ##variables
    ##Page1
    annexB_page1_1_1 = "This supplementary report is generated by default, even if the hospital admission data file is unavailable. The management of clinical and laboratory practice can be supported by some data indictors such as blood culture contamination rate, proportion of notifiable antibiotic-pathogen combinations, and proportion of isolates with infrequent phenotypes or potential errors in AST results. Isolates with infrequent phenotypes or potential errors in AST results include (a) reports of organisms which are intrinsically resistant to an antibiotic but are reported as susceptible and (b) reports of organisms with discordant AST results. " #3.1 3104
    annexB_page1_1_2 = "This supplementary report could support the clinicians, policy makers and the laboratory staff to understand their summary data quickly. The laboratory staff could also use \"Supplementary_data_indicators_report.pdf\" generated in the folder \"Report_with_patient_identifiers\" to check and validate individual data records further. "
    annexB_page1_1_3 = "<b>This supplementary report was estimated from data of blood specimens only.</b> Please note that the data indicators do not represent quality of the clinical or laboratory practice."
    annexB_page1_1 = [annexB_page1_1_1, 
                    add_blankline + annexB_page1_1_2, 
                    add_blankline + annexB_page1_1_3]

    annexB_page1_2_1 = "*Blood culture contamination rate is defined as the number of raw contaminated cultures per number of blood cultures received by the laboratory per reporting period. Blood culture contamination rate will not be estimated in case that the data of negative culture (specified as \"no growth\" in the dictionary file for microbiology data) is not available. " #3.1 3104
    annexB_page1_2_2 = "**Notifiable antibiotic-pathogen combinations and their classifications are defined as WHO list of AMR priority pathogen published in 2017. "
    annexB_page1_2_3 = "**, ***The proportion is estimated per number of blood specimens culture positive for any organisms with AST result in the raw microbiology data. "
    annexB_page1_2_4 = "*, **, ***Details of the criteria are available in Table 3 and Table 4 of \"Supplementary_data_indicators_report.pdf\", and \"list_of_indicators.xlsx\" in the folder \"Configuration\". "
    annexB_page1_2_5 = "NA = Not applicable"
    annexB_page1_2 = [annexB_page1_2_1+annexB_page1_2_2+annexB_page1_2_3+annexB_page1_2_4+annexB_page1_2_5]
    ##Page2
    annexB_page2_1_2 = "Data was stratified by month to assist detection of missing data and understand the change of indicators by months."
    annexB_page2_1 = [annexB_page2_1_2]
    ########### ANNEX B: PAGE1 ########
    ARC.report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Introduction',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,annexB_page1_1, 1.0*inch, 4.5*inch, 460, 350, font_size=11)
    ARC.report_title(c,'Results',1.07*inch, 5.2*inch,'#3e4444',font_size=12)
    table_draw = ARC.report_table_annexB_page1(blo_table)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.07*inch, 2.5*inch)
    ARC.report_context(c,annexB_page1_2, 1.0*inch, 0.6*inch, 460, 130, font_size=9, line_space=14)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ########### ANNEX B: PAGE2 ########
    ARC.report_title(c,"Annex B: Supplementary report on data indicators",1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_title(c,'Reporting period by months',1.07*inch, 9.5*inch,'#3e4444',font_size=12)
    ARC.report_context(c,annexB_page2_1, 1.0*inch, 8.7*inch, 460, 50, font_size=11)
    table_draw = ARC.report_table_annexB(blo_table_bymonth)
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.2*inch, 4.8*inch)
    ARC.report_context(c,annexB_page1_2, 1.0*inch, 0.8*inch, 460, 130, font_size=9,line_space=14)
    canvas_printpage(c,startpage+1,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    
def method(c,logger,lst_org_format,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    green_op = "<font color=darkgreen>"
    green_ed = "</font>"
    add_blankline = "<br/>"
    ##variables
    ##Page1
    method_page1_1 = "<b>" + "Data source:" + "</b>"
    method_page1_2 = "For each run (double−click on AMASS.bat file), AMASS application used the microbiology data file (microbiology_data) and the hospital admission data file (hospital_admission_data) that were stored in the same folder as the application file. " + \
                     "Hence, in case that the user would like to update, correct, revise or change the data, the data files in the folder should be updated before AMASS.bat file is double−clicked again. " + \
                     "A new report based on the updated data will then be generated."
    method_page1_3 = add_blankline + "<b>" + "Requirements:" + "</b>"
    method_page1_4 = "<b>" + "− Computer with Microsoft Windows 7 or higher" + "</b>"
    method_page1_5 = "AMASS may work in other versions of Microsoft Windows and other operating systems. " + \
                     "However, thorough testing and adjustment have not been performed."
    method_page1_6 = "<b>" + "− AMASS3.1.zip package file" + "</b>" #3.1 3104
    method_page1_7 = "AMASS application is to be downloaded from " + "<u><link href=\"https://www.amass.website\" color=\"blue\"fontName=\"Helvetica\">https://www.amass.website</link></u>" + ", and unzipped to generate an AMASS folder that could be stored under any folder in the computer. " + \
                    "AMASS folder contains 4 files (AMASS.bat, dictionary_for_microbiology_data.xlsx, dictionary_for_hospital_admission_data.xlsx, and dictionary_for_wards.xlsx), and 6 folders (Configuration, Example_dataset_1_WHONET, Example_dataset_2, Example_dataset_3_long_format, Example_dataset_4_cluster_signals, and Programs)." #3.1 3104
    method_page1_8 = "<b>" + "− Microbiology data file (microbiology data in .csv or .xlsx file format)" + "</b>" #3.1 3104
    method_page1_9 = "The user needs to obtain microbiology data, and then copy & paste this data file into the same folder as AMASS.bat file."
    method_page1_10 = "<b>" + "− [Optional] Hospital admission data file (hospital admission data in .csv or .xlsx file format)" + "</b>" #3.1 3104
    method_page1_11 = "If available, the user could obtain hospital admission data, and then copy & paste this data file into the same folder as AMASS.bat file."
    method_page1_12 = add_blankline + "<b>" + "Not required:" + "</b>"
    method_page1_13 = "<b>" + "− Internet to run AMASS application" + "</b>"
    method_page1_14 = "AMASS application will run offline. " + \
                      "No data are transferred while the application is running and reports are being generated. " + \
                      "The automatically generated reports are in PDF format (do not contain any patient identifier) and can be shared under the user's jurisdiction."
    method_page1 = [method_page1_1, method_page1_2, method_page1_3, method_page1_4, method_page1_5, 
                    method_page1_6, method_page1_7, method_page1_8, method_page1_9, method_page1_10, 
                    method_page1_11, method_page1_12, method_page1_13, method_page1_14]
    ##Page2
    method_page2_1_1 = "<b>" + "− Python" + "</b>"
    method_page2_1_2 = "The download package (AMASS3.1.zip) included Python portable and their libraries that AMASS application requires. " + \
                      "The user does not need to install any programme before using AMASS. " + \
                      "The user also does not have to uninstall Python if the computer already has the programme installed. " + \
                      "The user does not need to know how to use Python." #3.1 3104
    method_page2_1_3 = "<b>" + "− SaTScan" + "</b>"
    method_page2_1_4 = "The download package (AMASS3.1.zip) included batch SaTScan. " + \
                      "The user does not need to install SaTScan or any programme before using AMASS3.1. " + \
                      "The user does not need to know how to use SaTScan. " + \
                      "The user can configurate and edit the parameter values to run the cluster detection analyses through the file provided under the Configuration folder. " #3.1 3104
    method_page2_1 = [method_page2_1_1, method_page2_1_2, method_page2_1_3, method_page2_1_4]

    method_page2_2_1 = add_blankline + green_op + "<b>" + "Note:" + "</b>" + green_ed
    method_page2_2_2 = green_op + "[1] Please ensure that the file names of microbiology data file (microbiology_data) and the hospital admission data file (hospital_admission_data) are identical to what is written here. " + \
                       "Please make sure that all are lower−cases with an underscore \"_\" at each space." + green_ed #3.1 3104
    method_page2_2_3 = green_op + "[2] Please ensure that both microbiology and hospital admission data files have no empty rows. " + \
                       "For example, please do not add an empty row before the row of the variable names, which are the first row in both files)." + green_ed
    method_page2_2_4 = green_op + "[3] For the first run, a user may need to fill the data dictionary files to make sure that AMASS application understands your variable names and values." + green_ed
    method_page2_2 = [method_page2_2_1, method_page2_2_2, method_page2_2_3, method_page2_2_4]

    method_page2_3_1 = add_blankline + "AMASS uses a tier−based approach. " + \
                       "In cases when only the microbiology data file with the results of culture-negative specimens is not available, only section one, two, and three would be generated for users. " + \
                       "Section three would be generated only when data on admission date are available. " + \
                       "This is because these data are required for the stratification by origin of infection. " + \
                       "Section four would be generated only when data of specimens with culture negative (no microbial growth) are available in the microbiology data. " + \
                       "This is because these data are required for calculating the AMR frequency. " + \
                       "Section five would be generated only when both data of specimens with culture negative and admission date are available. " + \
                       "Section six would be generated only when mortality data are available."
    method_page2_3_2 = add_blankline + "Mortality was calculated from the number of in−hospital deaths (numerator) over the total number of patients with blood culture positive for the organism (denominator). " + \
                       "Please note that this is the all−cause mortality calculated using the outcome data in the data file, and may not necessarily represent the mortality directly due to the infections."
    method_page2_3 = [method_page2_3_1, method_page2_3_2]
    method_page2 = method_page2_1 + method_page2_2 + method_page2_3

    ##Page3
    method_page3_1_1 = "To detect spatio-temporal clusters of antimicrobial resistant bacterial species, AMASS-SaTScan used the retrospective space-time uniform model of the SaTScan (<u><link href=\"https://www.satscan.org\" color=\"blue\"fontName=\"Helvetica\">https://www.satscan.org</link></u>). " + \
                       "The cluster detection was based on the first hospital-origin resistant isolate per organism per patient per evaluation period. " + \
                       "Analyses were conducted separately for each of the seven species-groups, including MRSA, VREfs, VREfm, CREC, CRKP, CRPA, and CRAB identified from blood specimens only and from all types of specimens. " + \
                       "Both ward names (or ward identifiers) and resistant profiles were defined as \"location\" in the SaTScan to allow the detection of spatio-temporal cluster of periods with a higher than the expected frequency of a specific resistance profile. " + \
                       "AMASS-SaTScan assumed that each ward was independent. " + \
                       "In case that the ward name variable is not available (or some of the ward names are not filled in the dictionary file for wards), the whole hospital (or the wards that had no data in the dictionary files for wards) would be considered as a single space. " + \
                       "The total resistance isolates were used as the denominator. " + \
                       "Hypothesis testing was conducted using Monte Carlo simulations." #3.1 3104
    method_page3_1_2 = add_blankline + "<b>" + "How to use data dictionary files" + "</b>"
    method_page3_1_3 = "In cases when variable names in the microbiology and hospital admission data files were not the same as the one that AMASS used, the data dictionary files could be edited. " + \
                    "The raw microbiology and hospital admission data files were to be left unchanged. " + \
                    "The data dictionary files provided could be edited and re−used automatically when the microbiology and hospital admission data files were updated and AMASS.bat were to be double−clicked again (i.e. the data dictionary files would allow the user to re−analyze data files without the need to adjust variable names and data value again every time)."
    method_page3_1_4 = add_blankline + "For example:"
    method_page3_1_5 = "If variable name for \"hospital number\" is written as \"hn\" in the raw data file, the user would need to add \"hn\" in the cell next to \"hospital_number\". " + \
                    "If data value for blood specimens is defined by \"Blood−Hemoculture\" in the raw data file, then the user would need to add \"Blood−Hemoculture\" in the cell next to \"blood_specimen\"." #3.1 3104
    method_page3_1 = [method_page3_1_1, method_page3_1_2, method_page3_1_3, method_page3_1_4, method_page3_1_5]
    ##Page4
    method_page4_1 = ["<b>" + "Dictionary file (dictionary_for_microbiology_data.xlsx) may show up as in the table below:" + "</b>"]
    table_med_1 = [["Variable names used in AMASS", "Variable names used in \n your microbiology data file", "Requirements"],
                ["Don't change values in this \n column, but you can add rows \n with similar values if you need", 
                    "Change values in this column to \n represent how variable names \n are written in your raw \n microbiology data file", ""], 
                ["hospital_number", "", "Required"], 
                ["Values described in AMASS", "Values used in your \n microbiology data file", "Requirements"], 
                ["blood_specimen", "", "Required"]]
    method_page4_2 = ["<b>" + "Please fill in your variable names as follows:" + "</b>"]
    table_med_2 = [["Variable names used in AMASS", "Variable names used in \n your microbiology data file", "Requirements"],
                ["Don't change values in this \n column, but you can add rows \n with similar values if you need", 
                    "Change values in this column to \n represent how variable names \n are written in your raw \n microbiology data file", ""], 
                ["hospital_number", "hn", "Required"], 
                ["Values described in AMASS", "Values used in your \n microbiology data file", "Requirements"], 
                ["blood_specimen", "Blood−Hemoculture", "Required"]]
    method_page4_3 = ["Then, save the file. For every time the user double−clicked AMASS.bat, the application would know that the variable named \"hn\" is similar to \"hospital_number\" and represents the patient identifier in the analysis."] #3.1 3104
    
    ##Page5
    method_page5_1 = ["<b>" + "Organisms included for the AMR Surveillance Report:" + "</b>"] 
    method_page5_2 = []
    icountorg = len(lst_org_format)
    ihalf = m.ceil(icountorg/2)
    for i in range(icountorg):
        if i < ihalf:
            method_page5_1.append("− " + lst_org_format[i])
        else:
            method_page5_2.append("− " + lst_org_format[i])
    method_page5_1.append("The eight organisms and antibiotics included in the report were selected based on the global priority list of antibiotic resistant bacteria and Global Antimicrobial Resistance Surveillance System (GLASS) of WHO [1,2].")

    method_page5_5_1 = "<b>" + "Definitions:" + "</b>"
    method_page5_5_2 = "The definitions of infection origin proposed by the WHO GLASS was used [1]. In brief, community−origin bloodstream infection (BSI) was defined for patients in the hospital within the first two calendar days of admission when the first blood culture positive specimens were taken. " + \
                    "Hospital−origin BSI was defined for patients in the hospital longer than the first two calendar days of admission when the first blood culture positive specimens were taken. " + \
                    "In cases when the user had additional data on infection origin defined by infection control team or based on referral data, the user could edit the data dictionary file (variable name \"infection_origin\") and AMASS application would use the data of that variable to stratify the data by origin of infection instead of the above definition. " + \
                    "However, in cases when data on infection origin were not available (as in many hospitals in LMICs), the above definition would be calculated based on admission date and specimen collection date (with cutoff of 2 calendar days) and used to classify infections as community−origin or hospital−origin." #3.1 3104
    method_page5_5_3 = "<b>" + "De−duplication:" + "</b>"
    method_page5_5_4 = "When more than one blood culture was collected during patient management, duplicated findings of the same patient were excluded (de−duplicated). " + \
                    "Only one result was reported for each patient per sample type (blood) and surveyed organisms (listed above)." + \
                    "For example, if two blood cultures from the same patient had <i>E. coli</i>, only the first would be included in the report. " + \
                    "If there was growth of <i>E. coli</i> in one blood culture and of <i>K. pneumoniae</i> in the other blood culture, then both results would be reported. " + \
                    "One would be for the report on <i>E. coli</i> and the other one would be for the report on <i>K. pneumoniae</i>."
    method_page5_5 = [method_page5_5_1, method_page5_5_2, add_blankline + method_page5_5_3, method_page5_5_4]
    ##Backcover
    backcover_1_1 = "<b>" + "References:" + "</b>"
    backcover_1_2 = "[1] World Health Organization (2018) Global Antimicrobial Resistance Surveillance System (GLASS) Report. Early implantation 2016−2017. http://apps.who.int/iris/bitstream/handle/10665/259744/9789241513449−eng.pdf. (accessed on 3 Dec 2018)"
    backcover_1_3 = "[2] World Health Organization (2017) Global priority list of antibiotic−resistant bacteria to guide research, discovery, and development of new antibiotics. https://www.who.int/medicines/publications/WHO−PPL−Short_ Summary_25Feb−ET_NM_WHO.pdf. (accessed on 3 Dec 2018)"
    backcover_1 = [backcover_1_1, backcover_1_2, backcover_1_3]
    ########## METHOD: PAGE1 ##########
    ARC.report_title(c,'Methods used by AMASS application',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,method_page1, 1.0*inch, 0.7*inch, 460, 680, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ########## METHOD: PAGE2 ##########
    ARC.report_context(c,method_page2, 1.0*inch, 1.15*inch, 460, 700, font_size=11)
    # ARC.report_context(c,method_page2_1, 1.0*inch, 6.8*inch, 460, 300, font_size=11)
    # ARC.report_context(c,method_page2_2, 1.0*inch, 3.5*inch, 460, 300, font_size=11)
    # ARC.report_context(c,method_page2_3, 1.0*inch, 1.0*inch, 460, 300, font_size=11)
    # ARC.report_context(c,method_page3_1, 1.0*inch, 0.7*inch, 460, 170, font_size=11)
    canvas_printpage(c,startpage+1,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ########## METHOD: PAGE3 ##########
    ARC.report_context(c,method_page3_1, 1.0*inch, 2.4*inch, 460, 600, font_size=11)
    canvas_printpage(c,startpage+2,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage)
    ########## METHOD: PAGE4 ##########
    ARC.report_context(c,method_page4_1, 1.0*inch, 10.0*inch, 460, 50, font_size=11)
    table_draw = ARC.Table(table_med_1,  style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                                            ('FONT',(0,0),(2,0),'Helvetica-Bold'),
                                            ('FONT',(0,3),(-1,-2),'Helvetica-Bold'),
                                            ('FONTSIZE',(0,0),(-1,-1),11),
                                            ('TEXTCOLOR',(0,2),(-3,-3),colors.red),
                                            ('TEXTCOLOR',(0,4),(-3,-1),colors.red),
                                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                                            ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                            ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 8.0*inch)
    ARC.report_context(c,method_page4_2, 1.0*inch, 7.0*inch, 460, 50, font_size=11)
    table_draw = ARC.Table(table_med_2,  style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                                            ('FONT',(0,0),(2,0),'Helvetica-Bold'),
                                            ('FONT',(0,3),(-1,-2),'Helvetica-Bold'),
                                            ('FONTSIZE',(0,0),(-1,-1),11),
                                            ('TEXTCOLOR',(0,2),(-3,-3),colors.red),
                                            ('TEXTCOLOR',(0,4),(-3,-1),colors.red),
                                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                                            ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                            ('VALIGN',(0,0),(-1,-1),'MIDDLE')])
    table_draw.wrapOn(c, 500, 300)
    table_draw.drawOn(c, 1.0*inch, 5.2*inch)
    ARC.report_context(c,method_page4_3, 1.0*inch, 3.8*inch, 460, 80, font_size=11)
    # ARC.report_context(c,method_page5_1, 1.0*inch, 1.4*inch, 460, 180, font_size=11)
    # ARC.report_context(c,method_page5_2, 4.0*inch, 2.0*inch, 460, 120, font_size=11)
    canvas_printpage(c,startpage+3,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 
    ########## METHOD: PAGE5 ##########
    ARC.report_context(c,method_page5_1, 1.0*inch, 8.3*inch, 460, 180, font_size=11)
    ARC.report_context(c,method_page5_2, 4.0*inch, 8.9*inch, 460, 120, font_size=11)
    ARC.report_context(c,method_page5_5, 1.0*inch, 2.1*inch, 460, 450, font_size=11)
    u = inch/10.0
    c.setLineWidth(2)
    c.setStrokeColor(black)
    p = c.beginPath()
    p.moveTo(70,190) # start point (x,y)
    p.lineTo(7.45*inch,190) # end point (x,y)
    c.drawPath(p, stroke=1, fill=1)
    ARC.report_context(c,backcover_1, 1.0*inch, 0.5*inch, 460, 150, font_size=9)
    canvas_printpage(c,startpage+4,lastpage,today,False,ipagemode,ssectionanme,isecmaxpage,startpage) 

def investor(c,logger,startpage,lastpage, today=date.today().strftime("%d %b %Y"),ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1):
    ##paragraph variables
    add_blankline = "<br/>"
    ##variables
    ##Investor
    invest_1_1 = "AMASS application version 1.0 was developed by Cherry Lim, Clare Ling, Elizabeth Ashley, Paul Turner, Rahul Batra, Rogier van Doorn, Soawapak Hinjoy, Sopon Iamsirithaworn, Susanna Dunachie, Tri Wangrangsimakul, Viriya Hantrakun, William Schilling, John Stelling, Jonathan Edgeworth, Guy Thwaites, Nicholas PJ Day, Ben Cooper and Direk Limmathurotskul."
    invest_1_2 = "AMASS application version 2.0, 3.0 and 3.1 was developed by Chalida Rangsiwutisak, Preeyarach Klaytong, Prapass Wannapinij, Paul Tuner, John Stelling, Cherry Lim and Direk Limmathurotsakul."
    invest_1_3 = "AMASS application version 1.0 was funded by the Wellcome Trust (grant no. 206736 and 101103). C.L. was funded by a Research Training Fellowship (grant no. 206736) and D.L. was funded by an Intermediate Training Fellowship (grant no. 101103) from the Wellcome Trust."
    invest_1_4 = "AMASS application version 2.0, 3.0 and 3.1 was funded by the Wellcome Trust (grant no. 224681/Z/21/Z and Institutional Translational Partnership Award-MORU) "
    invest_1 = [invest_1_1, add_blankline+invest_1_2]
    
    invest_1_xtra = [invest_1_3, add_blankline+invest_1_4]
    
    invest_2_1 = "If you have any queries about AMASS, please contact:"
    invest_2_2 = "For technical information:"
    invest_2_3 = "Chalida Rangsiwutisak (chalida@tropmedes.ac),"
    invest_2_4 = "Cherry Lim (cherry@tropmedres.ac), and"
    invest_2_5 = "Direk Limmathurotsakul (direk@tropmedres.ac)"
    invest_2_6 = "For implementation of AMASS at your hospitals in Thailand:"
    invest_2_7 = "Preeyarach Klaytong (preeyarach@tropmedres.ac)"
    invest_2 = ["<b>"+invest_2_1+"</b>", 
                "<b>"+invest_2_2+"</b>", 
                invest_2_3, 
                invest_2_4, 
                invest_2_5, 
                "<b>"+add_blankline+invest_2_6+"</b>", 
                invest_2_7]
    ############# INVESTOR ############
    ARC.report_title(c,'Investigator team',1.07*inch, 10.5*inch,'#3e4444',font_size=16)
    ARC.report_context(c,invest_1, 1.0*inch, 6.0*inch, 460, 300, font_size=11)
    ARC.report_title(c,'Funding',1.07*inch, 7.0*inch,'#3e4444',font_size=16)
    ARC.report_context(c,invest_1_xtra, 1.0*inch, 2.5*inch, 460, 300, font_size=11)
    ARC.report_context(c,invest_2, 1.0*inch, 1.5*inch, 460, 200, font_size=11)
    canvas_printpage(c,startpage,lastpage,today,True,ipagemode,ssectionanme,isecmaxpage,startpage) 
    

# Main function called by AMR Analysis module -----------------------------------------------------------------------------------------------------------
def generate_amr_report(df_dict_micro,dict_orgcatwithatb,dict_orgwithatb_mortality,dict_orgwithatb_incidence,df_micro_ward,bishosp_ava,logger):
    AL.printlog("Start AMR Surveillance Report: " + str(datetime.now()),False,logger)
    sub_printprocmem("Start generate report",logger)
    AL.printlog("Import optional modules",False,logger)
    #If do annex C
    #dict_cmdarg  = AL.getcmdargtodict(logger)
    config=AL.readxlsxorcsv(path_input + "Configuration/", "Configuration",logger)
    bIsDoAnnexC = True
    try:
        #config=AL.readxlsxorcsv(AC.CONST_PATH_ROOT + "Configuration/", "Configuration",logger)
        bIsDoAnnexC = ARC.check_config(config, "amr_surveillance_annexC")
    except:
        pass
    if bIsDoAnnexC == True:
        try:
            import AMASS_annex_c_report as REP_ANNEX_C
            bIsDoAnnexC == True
            AL.printlog("Impoer Annex C report module",False,logger)
        except Exception as e:
            AL.printlog("Error : Import Annex C report module (AMASS_ANNEX_C_report.py) : " + str(e),True,logger)
            logger.exception(e)
            bIsDoAnnexC == False
            pass
    dict_ = df_dict_micro.iloc[:,:2].fillna("")
    dict_.columns = ["amass_name","user_name"]

    try:
        check_abaumannii = True if "organism_acinetobacter_baumannii" in list(dict_.loc[dict_["amass_name"]=="acinetobacter_spp_or_baumannii","user_name"]) else False
    except:
        check_abaumannii = False
    lst_org_format = []
    lst_org_full  = []
    lst_org_short = []
    #list of item for section 4,5,6 depending on interested organism in section 2, 3
    lst_org_rpt4_0 = []
    lst_org_rpt5_0 = []
    lst_org_rpt6_0 = []
    lst_org_short_rpt6_0 = []
    lst_org_format_rpt6_0 = []
    lst_atbbyorg_fromdict = []
    lst_itemperorg_fromdict_mor = []
    for sorgkey in dict_orgcatwithatb:
        ocurorg = dict_orgcatwithatb[sorgkey]
        #exclude no growth
        if ocurorg[1] == 1:
            lst_org_format.append(ocurorg[5])
            lst_org_full.append(ocurorg[2])
            lst_atbbyorg_fromdict.append(len(ocurorg[3]))
            
            org_1 = ocurorg[2].split(" ")        #['Staphylococcus], [aureus']   
            if 'spp' in org_1[1]:                #Staphylococcus spp -> sta_spp
                name = org_1[0][0:3]+"_spp"
            else:                                #Staphylococcus aureus -> s_aureus
                name = org_1[0][0]+"_"+org_1[1]
            lst_org_short.append(name.lower())  #['s_aureus', ...]
            if sorgkey in dict_orgwithatb_incidence.keys():
                ocurorg_incidence = dict_orgwithatb_incidence[sorgkey]
                #lst_org_rpt4_0.append(ocurorg_incidence[4])
                lst_org_rpt4_0.append(ocurorg_incidence[5])
                #list_atbname_incidence = ocurorg_incidence[1]
                list_atbname_incidence = ocurorg_incidence[6]
                list_atbname_incidence_mode = ocurorg_incidence[4]
                for iiatb in range(len(list_atbname_incidence)):
                    #Only list atb in mode defined (NS or RIS)
                    if list_atbname_incidence_mode[iiatb] == AC.CONST_MODE_R_OR_AST:
                        lst_org_rpt5_0.append(list_atbname_incidence[iiatb])
            if sorgkey in dict_orgwithatb_mortality.keys():
                ocurorg_mortalit = dict_orgwithatb_mortality[sorgkey]
                lst_org_rpt6_0.append(ocurorg_mortalit[0])
                lst_org_short_rpt6_0.append(name.lower())  
                lst_org_format_rpt6_0.append(ocurorg[5])
                lst_itemperorg_fromdict_mor.append(2*len(ocurorg_mortalit[1]))
    #dgendate = date.today()
    dgendate = datetime.now()
    strgendate = dgendate.strftime("%d %b %Y %H:%M")
    checkpoint_mic = ARC.checkpoint(path_input + "microbiology_data.xlsx") or ARC.checkpoint(path_input + "microbiology_data.csv")
    checkpoint_hos = ARC.checkpoint(path_input + "hospital_admission_data.xlsx") or ARC.checkpoint(path_input + "hospital_admission_data.csv")
    #config=AL.readxlsxorcsv(path_input + "Configuration/", "Configuration",logger)
    canvas_rpt = canvas.Canvas(path_input +"AMR_surveillance_report.pdf")
    ### Cover, genrate by -----------------------------------------------------------------------------------------------
    AL.printlog("AMR surveillance report - checkpoint cover",False,logger)
    try:
        df_cover_sec1_res = pd.read_csv(AC.CONST_PATH_RESULT + sec1_res_i).fillna("NA")
        cover(canvas_rpt,df_cover_sec1_res,strgendate)
        generatedby(canvas_rpt,df_cover_sec1_res)
    except Exception as e:
        AL.printlog("Error at : checkpoint print cover : " + str(e),True,logger)
        logger.exception(e)
        pass
    ### Table of content -------------------------------------------------------------------------------------------------
    ssecname_intro = 'Introduction'
    ssecname_sec1 = 'Section 1'
    ssecname_sec2 = 'Section 2'
    ssecname_sec3 = 'Section 3'
    ssecname_sec4 = 'Section 4'
    ssecname_sec5 = 'Section 5'
    ssecname_sec6 = 'Section 6'
    ssecname_annexA = 'Annex A'
    ssecname_annexB = 'Annex B'
    ssecname_annexC = 'Annex C'
    ssecname_method = 'Methods'
    ssecname_investor = 'Ackonwledgements'
    #Table is tricky it should be calculated upfront before generate other parts
    #Default number of page per section base on AMASS2.0
    ipage_intro = 2
    ipage_sec1 = 2
    ipage_sec2 = 7
    ipage_sec3 = 12
    ipage_sec4 = 3
    ipage_sec5 = 5
    ipage_sec6 = 6
    ipage_annexA1 = 2
    ipage_annexA2 = 1
    ipage_annexB = 2
    ipage_annexC = 0
    if bIsDoAnnexC == True:
        try:
            ipage_annexC = REP_ANNEX_C.get_annexC_roughtotalpage(logger)
        except:
            pass
    ipage_method = 5
    ipage_investor = 1
    #Cal page for sec2, 3 and 63 total page using to cal start page in sec 3, 4 and Annex A
    if AC.CONST_REPORTPAGENUM_MODE != 3: #mode 1 and 2 need to calculate total page per section (Mode-3 total page per section is fix regarding AMASS2.0)
        ipage_sec2 = 2
        ipage_sec3 = 2
        bIsStartofnewpage = True
        for i in range(len(lst_org_format)):
            bPrintPage = False
            if lst_atbbyorg_fromdict[i] > AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
                #Sec 2 full page, Sec 3 2 full page
                bPrintPage = True
                bIsStartofnewpage = True
                ipage_sec3 = ipage_sec3 + 2
            elif bIsStartofnewpage: 
                #Sec 2 top half, Sec 3 1 full page
                ii=i+1
                if ii>= len(lst_org_format):
                    bPrintPage = True
                else:
                    if lst_atbbyorg_fromdict[ii] <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
                        bPrintPage = False
                    else:
                        bPrintPage = True
                bIsStartofnewpage = False
                ipage_sec3 = ipage_sec3 + 1
            else:
                #Sec 2 bottom half, Sec 3 1 full page
                bPrintPage = True
                bIsStartofnewpage = True
                ipage_sec3 = ipage_sec3 + 1
            if bPrintPage:
                ipage_sec2 = ipage_sec2 + 1
                bIsStartofnewpage = True
        bIsStartofnewpage = True
        ipage_sec6 = 2
        for i in range(len(lst_org_rpt6_0)):
            bPrintPage = False
            if lst_itemperorg_fromdict_mor[i] > AC.CONST_MAX_ATBCOUNTFITHALFPAGE_MORALITY:
                bPrintPage = True
                bIsStartofnewpage = True
            elif bIsStartofnewpage: 
                ii = i+1
                if ii>= len(lst_org_rpt6_0):
                    bPrintPage = True
                else:
                    if lst_itemperorg_fromdict_mor[ii] <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE_MORALITY:
                        bPrintPage = False
                    else:
                        bPrintPage = True
                bIsStartofnewpage = False
            else:
                bPrintPage = True
                bIsStartofnewpage = True
            if bPrintPage:
                ipage_sec6 = ipage_sec6 + 1
                bIsStartofnewpage = True
    if AC.CONST_REPORTPAGENUM_MODE != 2: #Mode 1 and 3 is continue page number through out the report
        #Start cal start page number of each section base on total page (or max page in mode==3) of each section
        istartpage_intro = 1 
        istartpage_sec1 = istartpage_intro + ipage_intro 
        istartpage_sec2 = istartpage_sec1 + ipage_sec1     
        istartpage_sec3 = istartpage_sec2 + ipage_sec2  
        istartpage_sec4 = istartpage_sec3 + ipage_sec3 
        istartpage_sec5 = istartpage_sec4 + ipage_sec4
        istartpage_sec6 = istartpage_sec5 + ipage_sec5
        istartpage_annexA1 = istartpage_sec6 + ipage_sec6
        istartpage_annexA2 = istartpage_annexA1 + ipage_annexA1
        istartpage_annexB = istartpage_annexA2 + ipage_annexA2
        if bIsDoAnnexC == True:
            istartpage_annexC = istartpage_annexB + ipage_annexB
            istartpage_method = istartpage_annexC + ipage_annexC
        else:
            istartpage_method = istartpage_annexB + ipage_annexB
        istartpage_investor = istartpage_method + ipage_method
        ilastpage = istartpage_investor + ipage_investor
        ilastpage = ilastpage - 1
    else:
        istartpage_intro = 1 
        istartpage_sec1 = 1 
        istartpage_sec2 = 1    
        istartpage_sec3 = 1  
        istartpage_sec4 = 1 
        istartpage_sec5 = 1
        istartpage_sec6 = 1
        istartpage_annexA1 = 1
        istartpage_annexA2 = 1
        istartpage_annexB = 1
        istartpage_annexC = 1
        istartpage_method = 1
        istartpage_investor = 1
        ilastpage = 1
    #End page calculate and setting
    bgotdata_sec1 = ARC.checkpoint(AC.CONST_PATH_RESULT + sec1_num_i)
    bgotdata_sec2 = ARC.checkpoint(AC.CONST_PATH_RESULT +  sec2_res_i)
    bgotdata_sec3 = ARC.checkpoint(AC.CONST_PATH_RESULT +  sec3_res_i)
    bgotdata_sec4 = ARC.checkpoint(AC.CONST_PATH_RESULT +  sec4_res_i)
    bgotdata_sec5 = ARC.checkpoint(AC.CONST_PATH_RESULT +  sec5_res_i)
    bgotdata_sec6 = ARC.checkpoint(AC.CONST_PATH_RESULT +  sec6_res_i)
    bgotdata_annexA1 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secA_res_i)
    bgotdata_annexA11 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secA_res_i_A11)
    bgotdata_annexA2 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secA_mor_i)
    bgotdata_annexB_page1 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secB_blo_i)
    bgotdata_annexB_page2 = ARC.checkpoint(AC.CONST_PATH_RESULT +  secB_blo_mon_i)
    #print table of contents
    content_0 = 'Introduction'
    content_1 = 'Section [1]: Data overview'
    content_2 = 'Section [2]: AMR proportion report'
    content_3 = 'Section [3]: AMR proportion report with stratification by infection origin'
    content_4 = 'Section [4]: AMR frequency report'
    content_5 = 'Section [5]: AMR frequency report with stratification by infection origin'
    content_6 = 'Section [6]: Mortality involving AMR and antimicrobial−susceptible infections'
    content_7 = 'Annex A: Supplementary report on notifiable bacterial infections'
    content_8 = 'Annex B: Supplementary report on data indicators'
    content_opt = []
    istartpage_opt = []
    if bIsDoAnnexC == True:
        content_opt.append(REP_ANNEX_C.ANNEXC_RPT_CONST_TITLE)
        istartpage_opt.append(str(istartpage_annexC))
    content_9 = 'Methods'
    content_10 = 'Acknowledgements'
    content = [content_0, content_1, content_2, content_3, content_4, content_5, content_6, content_7, content_8] + content_opt + [content_9, content_10]
    content_page = [str(istartpage_intro), str(istartpage_sec1), str(istartpage_sec2), str(istartpage_sec3), str(istartpage_sec4), str(istartpage_sec5), str(istartpage_sec6), str(istartpage_annexA1), str(istartpage_annexB)]
    content_page = content_page + istartpage_opt
    content_page = content_page + [str(istartpage_method), str(istartpage_investor)]
    ############# CONTENT #############
    ARC.report_title(canvas_rpt,'Content',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
    ARC.report_context(canvas_rpt,content, 1.0*inch, 6.0*inch, 435, 300, font_size=11)
    if AC.CONST_REPORTPAGENUM_MODE != 2: 
        ARC.report_context(canvas_rpt,content_page, 7.0*inch, 6.0*inch, 30, 300, font_size=11)
    canvas_rpt.showPage()
    ### Introduction -----------------------------------------------------------------------------------------------------
    AL.printlog("AMR surveillance report - checkpoint introduction",False,logger)
    try:
        introduction(canvas_rpt,logger, istartpage_intro, ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_intro,ipage_intro)
    except Exception as e:
        AL.printlog("Error at : checkpoint print introduction : " + str(e),True,logger)
        logger.exception(e)
        pass 
    ### SECTION 1 -----------------------------------------------------------------------------------------------------
    if ARC.check_config(config, "amr_surveillance_section1"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 1",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 1",logger)
        if bgotdata_sec1:
            try:
                df_sec1_T = pd.read_csv(AC.CONST_PATH_RESULT + sec1_num_i)
                df_sec1_T = ARC.prepare_section1_table_for_reportlab(df_sec1_T,checkpoint_hos)
                section1(canvas_rpt,logger,bishosp_ava,df_cover_sec1_res,df_sec1_T,istartpage_sec1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec1,ipage_sec1)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 1 : " + str(e),True,logger)
                logger.exception(e)
                pass 
        else:
           AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 1",False,logger) 
           section1_nodata(canvas_rpt,istartpage_sec1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec1,ipage_sec1)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 1",False,logger) 
    ### SECTION 2 -----------------------------------------------------------------------------------------------------
    if ARC.check_config(config, "amr_surveillance_section2"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 2",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 2",logger)
        if bgotdata_sec2:
            try:
                sec2_res = pd.read_csv(path_result + sec2_res_i).fillna("NA")
                sec2_T_amr = pd.read_csv(path_result + sec2_amr_i)
                sec2_T_org = pd.read_csv(path_result + sec2_org_i)
                sec2_T_pat = pd.read_csv(path_result + sec2_pat_i)
                ##Section 2; Page 1
                sec2_merge = ARC.prepare_section2_table_for_reportlab_V3(sec2_T_org, sec2_T_pat, lst_org_full,lst_org_format, ARC.checkpoint(path_result + sec2_res_i),dict_orgcatwithatb)
                ##Section 2; Page 2-6
                #Retriving numper of positive patient of each organism
                #Creating AMR grapgh of each organism
                #Creating AMR table of each organism
                lst_numpat = []
                list_sec2_org_table = list()
                list_sec2_atbcountperorg = list()
                list_sec2_atbnote = list()
                for i in range(len(lst_org_full)):
                    numpat = ARC.create_num_patient(sec2_T_pat, lst_org_full[i], 'Organism')
                    lst_numpat.append(numpat)
                    
                    palette = ARC.create_graphpalette(sec2_T_amr,'Total(N)','Organism',lst_org_full[i],numpat,cutoff=70.0)
                    #Generate graph and saved to PNG files
                    iatbrow = ARC.count_atbperorg(sec2_T_amr,lst_org_full[i])
                    ifig_H = 2*ARC.cal_sec2and3_fig_height(iatbrow)
                    sec2_G = ARC.create_graph_nons_v3(sec2_T_amr,lst_org_full[i], lst_org_short[i],palette,'Antibiotic',AC.CONST_MODE_R_OR_AST,ifig_H)
                    list_sec2_org_table.append(ARC.create_table_nons_V3(sec2_T_amr,lst_org_full[i],"",AC.CONST_MODE_R_OR_AST).values.tolist())
                    list_sec2_atbcountperorg.append(iatbrow)
                    list_sec2_atbnote.append(ARC.get_atbnoteperorg(sec2_T_amr,lst_org_full[i]))
                #Print the graph
                section2(canvas_rpt,logger,sec2_res, sec2_merge, lst_org_format,lst_numpat, lst_org_short, list_sec2_org_table,list_sec2_atbcountperorg,list_sec2_atbnote,
                         istartpage_sec2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec2,ipage_sec2)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 2 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 2",False,logger) 
            section2_nodata(canvas_rpt,istartpage_sec2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec2,ipage_sec2)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 2",False,logger) 
    ### SECTION 3 -----------------------------------------------------------------------------------------------------
    if ARC.check_config(config, "amr_surveillance_section3"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 3",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 3",logger)
        if bgotdata_sec3:
            try:
                sec3_res = pd.read_csv(path_result + sec3_res_i).fillna("NA")
                sec3_amr = pd.read_csv(path_result + sec3_amr_i)
                sec3_pat = pd.read_csv(path_result + sec3_pat_i)
                sec3_pat_1 = sec3_pat.copy()
                ##Section 3; Page 1
                sec3_pat = pd.read_csv(path_result + sec3_pat_i).drop(columns=["Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file"])
                sec3_pat_sum = sec3_pat.copy().fillna(0)
                sec3_pat = sec3_pat.astype(str)
                sec3_pat_1 = sec3_pat.astype(str)
                sec3_pat.loc["Total","Organism"] = "Total"
                sec3_pat.loc["Total","Number_of_patients_with_blood_culture_positive"] = str(sec3_pat_sum["Number_of_patients_with_blood_culture_positive"].sum())
                sec3_pat.loc["Total","Community_origin"] = str(round(sec3_pat_sum["Community_origin"].sum()))
                sec3_pat.loc["Total","Hospital_origin"] = str(round(sec3_pat_sum["Hospital_origin"].sum()))
                sec3_pat.loc["Total","Unknown_origin"] = str(round(sec3_pat_sum["Unknown_origin"].sum()))
                sec3_pat = sec3_pat.rename(columns={"Number_of_patients_with_blood_culture_positive":"Number of patients with\nblood culture positive\nfor the organism", 
                                                    "Community_origin":"Community\n-origin**","Hospital_origin":"Hospital\n-origin**","Unknown_origin":"Unknown\n-origin***"})
                sec3_pat = ARC.prepare_section3_table_for_reportlab_V3(sec3_pat,lst_org_full,lst_org_format)
                sec3_col = pd.DataFrame(list(sec3_pat.columns),index=sec3_pat.columns).T
                sec3_pat_val = sec3_col.append(sec3_pat).values.tolist()

                ##Section 3; Page 3-12
                sec3_lst_numpat_CO = []  
                sec3_lst_numpat_HO = []  
                list_sec3_org_table_CO = []
                list_sec3_org_table_HO = []
                list_sec3_atbcountperorg = []
                list_sec3_atbnote_CO = []
                list_sec3_atbnote_HO = []
                for i in range(len(lst_org_full)):
                    iatbrow = ARC.count_atbperorg(sec3_amr,lst_org_full[i],"Hospital")
                    ifig_H = 2*ARC.cal_sec2and3_fig_height(iatbrow)
                    list_sec3_atbcountperorg.append(iatbrow)
                    for ori in ['Community_origin','Hospital_origin']:
                        numpat = ARC.create_num_patient(sec3_pat_1.astype(str), lst_org_full[i], 'Organism', ori)
                        palette = ARC.create_graphpalette(sec3_amr,'Total(N)','Organism',lst_org_full[i],float(numpat),cutoff=70.0,origin=ori[:-7])
                        sec3_G = ARC.create_graph_nons_v3(sec3_amr,lst_org_full[i], lst_org_short[i],palette,'Antibiotic',AC.CONST_MODE_R_OR_AST,ifig_H,origin=ori[:-7])
                        if ori == 'Community_origin':
                            sec3_lst_numpat_CO.append(numpat)
                            list_sec3_org_table_CO.append(ARC.create_table_nons_V3(sec3_amr,lst_org_full[i],origin='Community',iMODE=AC.CONST_MODE_R_OR_AST).values.tolist())
                            list_sec3_atbnote_CO.append(ARC.get_atbnoteperorg(sec3_amr,lst_org_full[i],origin='Community'))
                        else:
                            sec3_lst_numpat_HO.append(numpat)
                            list_sec3_org_table_HO.append(ARC.create_table_nons_V3(sec3_amr,lst_org_full[i],origin='Hospital',iMODE=AC.CONST_MODE_R_OR_AST).values.tolist())
                            list_sec3_atbnote_HO.append(ARC.get_atbnoteperorg(sec3_amr,lst_org_full[i],origin='Hospital'))
                #Print 
                section3(canvas_rpt,logger,sec3_res, sec3_pat_val,lst_org_format, sec3_lst_numpat_CO,sec3_lst_numpat_HO, lst_org_short,list_sec3_org_table_CO,list_sec3_org_table_HO,list_sec3_atbcountperorg,list_sec3_atbnote_CO,list_sec3_atbnote_HO,
                             istartpage_sec3,  ilastpage, strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec3,ipage_sec3)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 3 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 3",False,logger) 
            section3_nodata(canvas_rpt,istartpage_sec3,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec3,ipage_sec3)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 3",False,logger) 
    ##### section4 #####
    if ARC.check_config(config, "amr_surveillance_section4"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 4",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 4",logger)
        if bgotdata_sec4:
            try:
                sec4_res = pd.read_csv(path_result + sec4_res_i).fillna("NA")
                style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
                ##Section 4; Page 2
                lst_org_rpt4_graph = []
                lst_org_rpt4_table = []
                for i in range(len(lst_org_rpt4_0)): #page2
                    #lst_org_rpt4_graph.append(ARC.prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="graph"))
                    #lst_org_rpt4_table.append(Paragraph(ARC.prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="table"),style_summary))
                    lst_org_rpt4_graph.append(lst_org_rpt4_0[i].replace("<i>","$").replace("</i>","$"))
                    lst_org_rpt4_table.append(Paragraph(lst_org_rpt4_0[i],style_summary))                   
                sec4_blo = pd.read_csv(path_result + sec4_blo_i)
                sec4_blo_1 = ARC.create_table_surveillance_1(sec4_blo, lst_org_rpt4_table)
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec4_blo))
                ARC.create_graph_surveillance_V3(sec4_blo, lst_org_rpt4_graph, 'Report4_frequency_blood',ifig_H=ifig_H)
                ##Section 4; Page 3
                lst_org_rpt4_graph_page3 = []
                lst_org_rpt4_table_page3 = []
                for i in range(len(lst_org_rpt5_0)): #page3
                    #lst_org_rpt4_graph_page3.append(ARC.prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="graph", text_work_drug="Y"))
                    #lst_org_rpt4_table_page3.append(Paragraph(ARC.prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="table", text_work_drug="Y").replace(" \n<i>","<br/><i>"),style_summary))
                    lst_org_rpt4_graph_page3.append(lst_org_rpt5_0[i].replace("<i>","$").replace("</i>","$"))
                    lst_org_rpt4_table_page3.append(Paragraph(lst_org_rpt5_0[i].replace("\n","<br/>"),style_summary))
                sec4_pat = pd.read_csv(path_result + sec4_pri_i)
                #Fileter by AC.CONST_MODE_R_OR_AST which done during load dict 
                #sec4_pat = sec4_pat[sec4_pat['Priority_pathogen'].isin(lst_org_rpt5_0)]
                sec4_pat = sec4_pat[sec4_pat['IncludeonlyR'] == AC.CONST_MODE_R_OR_AST]
                #Reset index before concat in create_table_surveillance_1
                sec4_pat.reset_index(drop=True,inplace=True)
                sec4_pat_1 = ARC.create_table_surveillance_1(sec4_pat, lst_org_rpt4_table_page3, text_work_drug="Y")
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec4_pat))
                ARC.create_graph_surveillance_V3(sec4_pat, lst_org_rpt4_graph_page3, 'Report4_frequency_pathogen', text_work_drug="Y",ifig_H=ifig_H)
                section4(canvas_rpt,logger,sec4_res,sec4_blo_1,sec4_pat_1,sec4_pat,istartpage_sec4,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec4,ipage_sec4)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 4 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 4",False,logger) 
            section4_nodata(canvas_rpt,istartpage_sec4,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec4,ipage_sec4)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 4",False,logger) 
    ##### section5 #####
    if ARC.check_config(config, "amr_surveillance_section5"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 5",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 5",logger)
        if bgotdata_sec5:
            try:
                sec5_res = pd.read_csv(path_result + sec5_res_i).fillna("NA")
                style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
                lst_org_rpt5_graph = []
                lst_org_rpt5_table = []
                for i in range(len(lst_org_rpt4_0)): #page2
                    #lst_org_rpt5_graph.append(ARC.prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="graph"))
                    #lst_org_rpt5_table.append(Paragraph(ARC.prepare_org_core(lst_org_rpt4_0[i], text_line=1, text_style="short", text_work="table"),style_summary))
                    lst_org_rpt5_graph.append(lst_org_rpt4_0[i].replace("<i>","$").replace("</i>","$"))
                    lst_org_rpt5_table.append(Paragraph(lst_org_rpt4_0[i],style_summary))
                lst_org_rpt5_graph_page3 = []
                lst_org_rpt5_table_page3 = []
                for i in range(len(lst_org_rpt5_0)): #page3
                    #lst_org_rpt5_graph_page3.append(ARC.prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="graph", text_work_drug="Y"))
                    #lst_org_rpt5_table_page3.append(Paragraph(ARC.prepare_org_core(lst_org_rpt5_0[i], text_line=2, text_style="short", text_work="table", text_work_drug="Y").replace(" \n<i>","<br/><i>"),style_summary))
                    lst_org_rpt5_graph_page3.append(lst_org_rpt5_0[i].replace("<i>","$").replace("</i>","$"))
                    lst_org_rpt5_table_page3.append(Paragraph(lst_org_rpt5_0[i].replace("\n","<br/>"),style_summary))
                ##Section 5; Page 2
                sec5_com = pd.read_csv(path_result + sec5_com_i)
                sec5_com_1 = ARC.create_table_surveillance_1(sec5_com, lst_org_rpt5_table)
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec5_com))
                ARC.create_graph_surveillance_V3(sec5_com, lst_org_rpt5_graph, 'Report5_incidence_community',ifig_H=ifig_H)
                ##Section 5; Page 3
                sec5_hos = pd.read_csv(path_result + sec5_hos_i)
                sec5_hos_1 = ARC.create_table_surveillance_1(sec5_hos, lst_org_rpt5_table)
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec5_hos))
                ARC.create_graph_surveillance_V3(sec5_hos, lst_org_rpt5_graph, 'Report5_incidence_hospital',ifig_H=ifig_H)
                ##Section 5; Page 4
                sec5_com_amr = pd.read_csv(path_result + sec5_com_amr_i)
                #Fileter by AC.CONST_MODE_R_OR_AST which done during load dict 
                #sec5_com_amr = sec5_com_amr[sec5_com_amr['Priority_pathogen'].isin(lst_org_rpt5_0)]
                sec5_com_amr = sec5_com_amr[sec5_com_amr['IncludeonlyR'] == AC.CONST_MODE_R_OR_AST]
                #Reset index before concat in create_table_surveillance_1
                sec5_com_amr.reset_index(drop=True,inplace=True)
                sec5_com_amr_1 = ARC.create_table_surveillance_1(sec5_com_amr, lst_org_rpt5_table_page3)
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec5_com_amr))
                ARC.create_graph_surveillance_V3(sec5_com_amr, lst_org_rpt5_graph_page3, 'Report5_incidence_community_antibiotic', text_work_drug="Y",ifig_H=ifig_H)
                ##Section 5; Page 5
                sec5_hos_amr = pd.read_csv(path_result + sec5_hos_amr_i)
                #Fileter by AC.CONST_MODE_R_OR_AST which done during load dict 
                #sec5_hos_amr = sec5_hos_amr[sec5_hos_amr['Priority_pathogen'].isin(lst_org_rpt5_0)]
                sec5_hos_amr = sec5_hos_amr[sec5_hos_amr['IncludeonlyR'] == AC.CONST_MODE_R_OR_AST]
                #Reset index before concat in create_table_surveillance_1
                sec5_hos_amr.reset_index(drop=True,inplace=True)
                sec5_hos_amr_1 = ARC.create_table_surveillance_1(sec5_hos_amr, lst_org_rpt5_table_page3)
                ifig_H = 2*ARC.cal_sec4_fig_height(len(sec5_hos_amr))
                ARC.create_graph_surveillance_V3(sec5_hos_amr, lst_org_rpt5_graph_page3, 'Report5_incidence_hospital_antibiotic', text_work_drug="Y",ifig_H=ifig_H)
                section5(canvas_rpt,logger,sec5_res, sec5_com_1, sec5_hos_1, sec5_com_amr_1, sec5_hos_amr_1,sec5_hos_amr,
                         istartpage_sec5,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec5,ipage_sec5)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 5 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 5",False,logger) 
            section5_nodata(canvas_rpt,istartpage_sec5,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec5,ipage_sec5)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 5",False,logger) 
    ##### section6 #####
    if ARC.check_config(config, "amr_surveillance_section6"):
        #AL.printlog("AMR surveillance report - checkpoint SECTION 6",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint SECTION 6",logger)
        if bgotdata_sec6:
            try:
                sec6_res = pd.read_csv(path_result + sec6_res_i).fillna("NA")
                sec6_mor = pd.read_csv(path_result + sec6_mor_i)
                #Fileter by AC.CONST_MODE_R_OR_AST which done during load dict 
                sec6_mor = sec6_mor[sec6_mor['IncludeonlyR'] == AC.CONST_MODE_R_OR_AST]
                sec6_mor_byorg = pd.read_csv(path_result + sec6_mor_byorg_i).fillna(0)
                ##Creating table for page33
                sec6_mor_all = ARC.prepare_section6_mortality_table_for_reportlab_V3(sec6_mor_byorg,lst_org_full,lst_org_format)
                #### section6; page2-5 #####
                sec6_mor_1 = ARC.prepare_section6_mortality_table(sec6_mor)
                sec6_G = sec6_mor.copy().replace(regex=["\r"],value="")
                sec6_G["Mortality (n)"] = round(sec6_G["Number_of_deaths"]/sec6_G["Total_number_of_patients"]*100,1).fillna(0)
                list_sec6_mor_tbl_CO = []
                list_sec6_mor_tbl_HO = []
                for i in range(len(lst_org_rpt6_0)):
                    for ori in ['Community-origin','Hospital-origin']:
                        tmp_mor_tbl =  ARC.create_table_mortal(sec6_mor_1, lst_org_rpt6_0[i], ori).replace(regex=["\r"],value="").values.tolist()
                        ARC.create_graph_mortal_V3(sec6_G,lst_org_rpt6_0[i],ori,'Report6_mortality_'+lst_org_short_rpt6_0[i]+'_'+ori[:-7].lower())
                        if ori == 'Community-origin':
                            list_sec6_mor_tbl_CO.append(tmp_mor_tbl)
                        else:
                            list_sec6_mor_tbl_HO.append(tmp_mor_tbl)
                #number of patient
                sec6_numpat_com = ARC.prepare_section6_numpat_dict(sec6_mor_byorg, "Community-origin")
                sec6_numpat_hos = ARC.prepare_section6_numpat_dict(sec6_mor_byorg, "Hospital-origin")
                section6(canvas_rpt,logger,sec6_res,sec6_mor, sec6_mor_all, list_sec6_mor_tbl_CO,list_sec6_mor_tbl_HO,
                         lst_org_format_rpt6_0,lst_org_short_rpt6_0,lst_org_rpt6_0,sec6_numpat_com, sec6_numpat_hos,
                         istartpage_sec6,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec6,ipage_sec6)
            except Exception as e:
                AL.printlog("Error at : checkpoint print SECTION 6 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            AL.printlog("WARNING : AMR surveillance report - no analysis data for SECTION 6",False,logger) 
            section6_nodata(canvas_rpt,istartpage_sec6,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec6,ipage_sec6)
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled SECTION 6",False,logger)   
    ##### AnnexA #####
    if ARC.check_config(config, "amr_surveillance_annexA"):
        #AL.printlog("AMR surveillance report - checkpoint annex A,A1",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint annex A,A1",logger)
        lst_annexA_org_full = []
        lst_annexA_org_format= []
        lst_annexA_org_format_full= []
        for sorgkey in AC.dict_annex_a_listorg:
            ocurorg = AC.dict_annex_a_listorg[sorgkey]
            #exclude no growth
            if ocurorg[1] == 1:
                lst_annexA_org_format.append(ocurorg[5])
                try:
                    lst_annexA_org_format_full.append(ocurorg[6])
                except:
                    lst_annexA_org_format_full.append(ocurorg[5])
                lst_annexA_org_full.append(ocurorg[2])
        style_normal = ParagraphStyle('normal',fontName='Helvetica',fontSize=11,alignment=TA_LEFT)
        style_small = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
        if bgotdata_annexA1:
            try:
                secA_res = pd.read_csv(path_result + secA_res_i).fillna("NA")
                secA_pat = pd.read_csv(path_result + secA_pat_i).fillna("NA")
                lst_page1_l = []
                lst_page1_r = []
                lst_page2 = []
                icountorg = len(lst_annexA_org_format)
                ihalf = m.ceil(icountorg/2)
                for i in range(icountorg):
                    if i < ihalf:
                        lst_page1_l.append(Paragraph("- " + lst_annexA_org_format_full[i],style_normal))
                    else:
                        lst_page1_r.append(Paragraph("- " + lst_annexA_org_format_full[i],style_normal))
                    lst_page2.append(Paragraph(lst_annexA_org_format[i],style_small)) 
                annexA_org_page1 = []
                for i in range(len(lst_page1_l)):
                    sl = lst_page1_l[i] if (i<len(lst_page1_l)) else ""
                    sr = lst_page1_r[i] if (i<len(lst_page1_r)) else ""
                    annexA_org_page1.append([sl,sr])
                    
                #Creating formatted organism for AnnexB's table on page2
                secA_pat2 = ARC.prepare_annexA_numpat_table_for_reportlab(secA_pat, lst_page2)
                #Version 4 add Annex A1.1 ------------------
                secA_res_A11 = pd.DataFrame()
                secA_pat2_A11 = []
                if bgotdata_annexA11:
                    secA_res_A11 = pd.read_csv(path_result + secA_res_i_A11).fillna("NA")
                    secA_pat_A11 = pd.read_csv(path_result + secA_pat_i_A11).fillna("NA")
                    secA_pat2_A11 = ARC.prepare_annexA_numpat_table_for_reportlab(secA_pat_A11, lst_page2)
                #-------------------------------------------
                annexA_A1(canvas_rpt,logger,secA_res, annexA_org_page1, secA_pat2,secA_res_A11,  secA_pat2_A11,
                              istartpage_annexA1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1)
                #annexA_A1(canvas_rpt,logger,secA_res, annexA_org_page1, secA_pat2,secA_res_A11,  secA_pat2_A11,
                #              istartpage_annexA1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA2)

            except Exception as e:
                AL.printlog("Error at : checkpoint print ANNEX A, A1a, a1b : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            annexA_nodata(canvas_rpt,istartpage_annexA1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1)
            #annexA_nodata(canvas_rpt,istartpage_annexA1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA2)
            pass
        #AL.printlog("AMR surveillance report - checkpoint annex A2",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint annex A,A2",logger)
        if bgotdata_annexA2:
            try:
                ##Creating table for page40 (Parsing)
                secA_mortal = pd.read_csv(path_result + secA_mor_i).fillna(0)
                #Preparing organism for AnnexA page3
                style = ParagraphStyle('normal',fontName='Helvetica',fontSize=9,alignment=TA_LEFT)
                #annexA_org = secA_mortal["Organism"].tolist()
                lst_page3_table = []
                lst_page3_graph = []
                for i in range(len(lst_annexA_org_format)):
                    lst_page3_table.append(Paragraph(lst_annexA_org_format[i],style))
                    lst_page3_graph.append(lst_annexA_org_format[i].replace("<i>","$").replace("</i>","$"))
                #Creating formatted organism for AnnexB's table on page3
                annexA_org_page3 = pd.DataFrame(lst_page3_table,columns=["Organism_fmt"])
                ##Creating table for AnnexB page3
                secA_mortal3 = ARC.prepare_annexA_mortality_table_for_reportlab(secA_mortal, annexA_org_page3)
                ##Creating graph for AnnexB page3
                #sub_printprocmem("Before create graph",logger)
                ARC.create_annexA_mortality_graph(secA_mortal,lst_page3_graph)
                #sub_printprocmem("after create graph",logger)
                annexA_A2(canvas_rpt,logger,secA_mortal3,istartpage_annexA2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA2)
                #annexA_A2(canvas_rpt,logger,secA_mortal3,istartpage_annexA2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA2)
                #sub_printprocmem("afte put graph on page",logger)
            except Exception as e:
                AL.printlog("Error at : checkpoint print ANNEX A2 : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            annexA_nodata_mortality(canvas_rpt,istartpage_annexA2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA2)
            pass
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled ANNEX A",False,logger) 
    ##### AnnexB #####
    if ARC.check_config(config, "amr_surveillance_annexB"):
        #AL.printlog("AMR surveillance report - checkpoint annex B",False,logger)
        sub_printprocmem("AMR surveillance report - checkpoint annex B",logger)
        secB_blo_1 = []
        secB_blo_bymonth = []
        ##Creating table for AnnexB page1 (Parsing)
        if bgotdata_annexB_page1:
            try:
                secB_blo = pd.read_csv(path_result + secB_blo_i) #.fillna("NA")
                secB_blo_1 = ARC.prepare_annexB_summary_table_for_reportlab(secB_blo)
                ##Creating table for AnnexB page2 (Parsing)
                if bgotdata_annexB_page2:
                    try:
                        secB_blo_bymonth = pd.read_csv(path_result + secB_blo_mon_i) #.fillna("NA")
                        secB_blo_bymonth = ARC.prepare_annexB_summary_table_bymonth_for_reportlab(secB_blo_bymonth)
                    except Exception as e:
                        AL.printlog("Error at : checkpoint print ANNEX B (Page 2) : " + str(e),True,logger)
                        logger.exception(e)
                        pass
            except Exception as e:
                AL.printlog("Error at : checkpoint print ANNEX B (Page 1) : " + str(e),True,logger)
                logger.exception(e)
                pass
            try:
                annexB(canvas_rpt,logger,secB_blo_1, secB_blo_bymonth,istartpage_annexB,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexB,ipage_annexB)
            except Exception as e:
                AL.printlog("Error at : checkpoint print ANNEX B (Print page) : " + str(e),True,logger)
                logger.exception(e)
                pass
        else:
            annexB_nodata(canvas_rpt,istartpage_annexB,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexB,ipage_annexB)
            pass
    else:
        AL.printlog("WARNING : AMR surveillance report - Disabled ANNEX B",False,logger) 
    #If do annex C
    if bIsDoAnnexC == True:
        try:
            REP_ANNEX_C.generate_annex_c_report(canvas_rpt,logger,df_micro_ward,istartpage_annexC,ilastpage,ipage_annexC, strgendate,bishosp_ava)
        except Exception as e:
            AL.printlog("Error at : checkpoint print ANNEX C (Generate report) : " + str(e),True,logger)
            logger.exception(e)
            pass
    else:
        pass
    #Last 2 page
    method(canvas_rpt,logger,lst_org_format,istartpage_method,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_method,ipage_method)
    investor(canvas_rpt,logger,istartpage_investor,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_investor,ipage_investor)
    #For testing remove when production
    """
    section1_nodata(canvas_rpt,istartpage_sec1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec1,ipage_sec1)
    section2_nodata(canvas_rpt,istartpage_sec2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec2,ipage_sec2)
    section3_nodata(canvas_rpt,istartpage_sec3,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec3,ipage_sec3)
    section4_nodata(canvas_rpt,istartpage_sec4,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec4,ipage_sec4)
    section5_nodata(canvas_rpt,istartpage_sec5,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec5,ipage_sec5)
    section6_nodata(canvas_rpt,istartpage_sec6,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_sec6,ipage_sec6)
    annexA_nodata(canvas_rpt,istartpage_annexA1,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA1)
    annexA_nodata_mortality(canvas_rpt,istartpage_annexA2,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexA,ipage_annexA1+ipage_annexA1)
    annexB_nodata(canvas_rpt,istartpage_annexB,ilastpage,strgendate,AC.CONST_REPORTPAGENUM_MODE,ssecname_annexB,ipage_annexB)
    """
    canvas_rpt.save()
