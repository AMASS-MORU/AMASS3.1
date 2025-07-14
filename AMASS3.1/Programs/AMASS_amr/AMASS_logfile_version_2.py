#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.1 (AMASS version 3.1) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Data verification logfile reports systematically.

# Created on 20 APR 2022
# Last update on: 15 JUL 2025 #3.1 3104
import logging #for creating logfile
import pandas as pd #for creating and manipulating dataframe
import math
from datetime import date #for generating today date
import datetime 
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.platypus import * #for plotting graph and tables
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
#from AMASS_logfile_function_version_2 import * #for importing logfile functions
from reportlab.lib.colors import * #for importing color palette
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
import AMASS_logfile_function_version_2 as ALogL

#Global variables
CONST_DVL_TITLE_H_INCH = 0.5
CONST_DVL_TBLHEAD_H_INCH = 2.0
CONST_DVL_ROW_H_INCH = 0.25
CONST_DVL_MAX_PAGE_H_INCH = 10.0
CONST_DVL_MAXROW_PER_PAGE = math.floor((CONST_DVL_MAX_PAGE_H_INCH-(CONST_DVL_TBLHEAD_H_INCH + CONST_DVL_TITLE_H_INCH))/CONST_DVL_ROW_H_INCH)
def multi_ts_V3(c,logger,list_df,list_title,list_sub,list_bnewpage_after) :
    bprintpage_after = False
    itop = CONST_DVL_MAX_PAGE_H_INCH
    ispaceleft = 0
    for i in range(len(list_df)):
        ts_v3(c,logger,list_df[i],list_title[i],list_sub[i],itop,list_bnewpage_after[i],ispaceleft)
        bprintpage_after = list_bnewpage_after[i]
        itop = ispaceleft
    #For last page of the group
    if bprintpage_after == False:
        c.showPage()
def ts_v3(c,logger,df,title_name,title_sub,itop,bnewpage_after,ispaceleft_after) :
    if (itop - (CONST_DVL_TBLHEAD_H_INCH + CONST_DVL_TITLE_H_INCH)) <=(3*CONST_DVL_ROW_H_INCH) : #if space left from last section can contains least than 3 records.
        #End page of previous section
        c.showPage()
        itop = CONST_DVL_MAX_PAGE_H_INCH
    icurloc = 0
    istartloc = 0
    bisfirstpage = True
    while icurloc < len(df):
        if icurloc > 0:
            #End page of previous set
            c.showPage()
            itop = CONST_DVL_MAX_PAGE_H_INCH
        istartloc = icurloc
        inrec_Inch = itop - (CONST_DVL_TBLHEAD_H_INCH + CONST_DVL_TITLE_H_INCH)
        inrec = math.floor(inrec_Inch/CONST_DVL_ROW_H_INCH)
        if (istartloc + inrec) >= len(df) :
            inrec = len(df) - istartloc
        iendloc = icurloc + inrec -1 
        if iendloc > len(df) - 1:
            inrec = len(df) - istartloc
            iendloc = len(df) - 1
        temp_df = df.loc[istartloc:iendloc,]
        #Print
        stitle = "<b>" + title_name + " (continue): " + title_sub + "</b>"
        if bisfirstpage:
            stitle = "<b>" + title_name + ": " + title_sub + "</b>"           
        ALogL.report_context(c,[stitle], 1.0*inch, itop*inch, 460, 80, font_size=11)
        table_draw = ALogL.report_table_appendix(temp_df)
        table_draw.wrapOn(c, 500, 300)
        h = (CONST_DVL_MAXROW_PER_PAGE-len(temp_df))*(CONST_DVL_ROW_H_INCH) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
        table_draw.drawOn(c, 1.07*inch, (h+CONST_DVL_TITLE_H_INCH)*inch)
        icurloc = iendloc + 1
        bisfirstpage = False
    if bnewpage_after:
        c.showPage()
        ispaceleft_after = 0
    else:
        inrec_lastpage = len(df) - istartloc
        ispaceleft = (inrec_lastpage*CONST_DVL_ROW_H_INCH) + CONST_DVL_TBLHEAD_H_INCH + CONST_DVL_TITLE_H_INCH
        #Leave end dicision to next section or end group

def indent(sstr,ilevel):
    ##paragraph variables
    iden1_op = "<para leftindent=\"25\">"
    iden2_op = "<para leftindent=\"50\">"
    iden3_op = "<para leftindent=\"75\">"
    listiden = [iden1_op,iden2_op,iden3_op]
    iden_ed = "</para>"
    iden_s =iden1_op
    try:
        iden_s = listiden[ilevel-1]
    except:
        iden_s =iden1_op 
    return iden_s+ sstr + iden_ed
def indent1(sstr):
    return indent(sstr,1)
def indent2(sstr):
    return indent(sstr,2)
def indent3(sstr):
    return indent(sstr,3)
def boldstr(sstr):
    return "<b>"+ sstr + "</b>"
def fn_dateexample(df_date,sdatename,logger):
    #date field
    scol_dateformat = "dateformat"
    scol_example = "exampledate"
    list_sd = []
    #sd = ""
    try:
        if len(df_date) <= 0:
            #sd = indent2("Unable to detect " + sdatename.lower() +" format.")
            list_sd = [indent2("Unable to detect " + sdatename.lower() +" format.")]
            AL.printlog("Unable to detect " + sdatename.lower() +" format. (Log file not available or no data)",False,logger)
        elif len(df_date) == 1:
            #sd = indent2(sdatename +" are in " + boldstr(df_date.iloc[0][scol_dateformat]) + " format. Example : " + boldstr(df_date.iloc[0][scol_example]))
            list_sd = [indent2(sdatename +" are in " + boldstr(df_date.iloc[0][scol_dateformat]) + " format. Example : " + boldstr(df_date.iloc[0][scol_example]))]
        else:
            #sd = indent2(sdatename +" are in multiple formats.")
            list_sd = [indent2(sdatename +" are in multiple formats.")]
            for index, row in df_date.iterrows():
                #sd = sd + indent3(boldstr(row[scol_dateformat]) + " format. Example : " + boldstr(row[scol_example]))
                list_sd = list_sd + [indent3(boldstr(row[scol_dateformat]) + " format. Example : " + boldstr(row[scol_example]))]
        pass
    except Exception as e:
        AL.printlog("Unable to detect " + sdatename.lower() +" format.",True,logger)
        logger.exception(e)
        #sd = indent2("Unable to detect " + sdatename.lower() +" format.")
        list_sd = [indent2("Unable to detect " + sdatename.lower() +" format.")]
    #return sd
    return list_sd
    
def fn_hnexample(df_hn,shnname,logger):
    #date field
    scol_length = "str_length"
    scol_num = "counts"
    scol_per = "percent"
    list_shn = []
    try:
        df_hn = df_hn.sort_values(by=scol_num,ascending=False)
        #print(df_hn)
        if len(df_hn) <= 0:
            list_shn = [indent2("Unable to detect " + shnname.lower() + " charecter lengths.")]
            AL.printlog("Unable to detect " + shnname.lower() + " charecter lengths. (Log file not available or no data)",False,logger)
        elif len(df_hn) == 1:
            list_shn = [indent2(shnname + " are all " + boldstr(str(df_hn.iloc[0][scol_length]))+ " charecters long.")]
        else:
            iAll = df_hn[scol_num].sum()
            sminl = str(df_hn[scol_length].min())
            smaxl = str(df_hn[scol_length].max())
            list_shn = [indent2(shnname + " are from " + boldstr(str(sminl))+ " to " + boldstr(str(smaxl)) + " charecter lengths. The following are most common lengths.")]
            i = 1
            for index, row in df_hn.iterrows():
                if i > 3:
                    break
                list_shn = list_shn + [indent3(boldstr(str(int(row[scol_length]))) +" charecters long : " + boldstr(str(round(row[scol_per],2)) + "% (" + str(int(row[scol_num])) + "/" + str(iAll) + ")"))]
                i = i + 1
        pass
    except Exception as e:
        AL.printlog("Unable to detect " + shnname.lower() + " charecter lengths.",True,logger)
        logger.exception(e)
        list_shn = [indent2("Unable to detect " + shnname.lower() + " charecter lengths.")]
    return list_shn
def cover(c,over_raw, today=date.today().strftime("%d %b %Y")):
    
    ##paragraph variable
    bold_blue_op = "<b><font color=\"#000080\">"
    bold_blue_ed = "</font></b>"
    add_blankline = "<br/>"
    hospital_name = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Hospital_name","Parameters"),coverpage=True)
    country_name  = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Country","Parameters"),coverpage=True)
    spc_date_start_cov = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Minimum_date","Parameters"),coverpage=True)
    spc_date_end_cov   = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Maximum_date","Parameters"),coverpage=True)    ##content
    
    
    
    cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
    cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
    cover_1_3 = "<b>Data from:</b>"
    cover_1_4 = bold_blue_op + str(spc_date_start_cov) + " to " + str(spc_date_end_cov) + bold_blue_ed
    cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
    
    cover_2_1 = "This report is for users to review variable names and data values used by microbiology data file and hospital admission data file saved within the same folder as the application file (AMASS.bat). This report can be used to assist users while completing the data dictionaries for both data files." #3.1 3104
    today = datetime.datetime.now().strftime("%d %b %Y %H:%M:%S")
    cover_2_2 = "<b>Generated on:</b>  " + bold_blue_op + today + bold_blue_ed
    cover_2_3 = "<b>Software version:</b>  " + bold_blue_op + AC.CONST_SOFTWARE_VERSION + bold_blue_ed
    cover_2 = [cover_2_1,cover_2_2,cover_2_3]
    ##reportlab
    c.setFillColor('#FCBB42')
    c.rect(0,590,800,20, fill=True, stroke=False)
    c.setFillColor(grey)
    c.rect(0,420,800,150, fill=True, stroke=False)
    ALogL.report_title(c,'Data verification log file',0.7*inch, 485,'white',font_size=28)
    ALogL.report_context(c,cover_1, 0.7*inch, 2.2*inch, 460, 240, font_size=18,line_space=26)
    ALogL.report_context(c,cover_2, 0.7*inch, 0.8*inch, 460, 120, font_size=11)
    c.showPage()


def ts0(c,logger,over_raw,fspcdate,fadmdate,fdisdate,fmhn,fhhn,bhosp_avi):
    try:
        
        #Micro data
        #read dates
        df_spcdate = pd.read_excel(fspcdate)
        df_hn = pd.read_excel(fmhn) 
        log_1_1 = "<br/>Please review the following information carefully before interpreting the AMR surveillance report generated by AMASS application."
        log_1_2 = "<br/>The microbiology data file (stored in the same folder as the application file) had:" #3.1 3104
        sMicroNum = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_records","Parameters"))
        sMicroDatefrom = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Minimum_date","Parameters"))
        sMicroDateto = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Maximum_date","Parameters"))
        log_1_3 = indent1(boldstr(sMicroNum) + " records with collection dates ranging from " + boldstr(sMicroDatefrom)+ " to " + boldstr(sMicroDateto))
        #log_1_4 = fn_dateexample(df_spcdate,"Collection date",logger)
        list_log_1_4 = fn_dateexample(df_spcdate,"Collection date",logger)
        list_log_1_5 = fn_hnexample(df_hn,"Microbiology data's hospital number",logger)
        list_log_1_5_1 = ""
        try:
            if int(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_HN_with_leadingzero","Parameters")) > 0:
                list_log_1_5_1 = indent3("hospital number are some with leading zero.")
        except:
            pass
        log_1_6 = indent2("Missing data.")
        log_1_7 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_or_unknown_specimen_date","Parameters"))) + " records are missing or unknown format of collection date.")
        log_1_7_1 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_wrong_specimen_date","Parameters"))) + " records are wrong date or wrong date format of collection date.")
        log_1_8 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_specimen_type","Parameters"))) + " records are missing specimen type.")
        log_1_9 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_missing_culture_result","Parameters"))) + " records are missing culture result.")
        #log_1 = [log_1_1,log_1_2,log_1_3,log_1_4,log_1_5,log_1_6,log_1_7,log_1_8,log_1_9]
        log_1 = [log_1_1,log_1_2,log_1_3] + list_log_1_4 + list_log_1_5 + [list_log_1_5_1] + [log_1_6,log_1_7,log_1_7_1,log_1_8,log_1_9]
        #Hosp data , merged data
        log_2 = ['<br/>No hospital admission data file found in the same folder as the application file. Thus, no hsopital admission data set information.'] #3.1 3104
        log_3 = ['<br/>No hospital admission data file found in the same folder as the application file. Thus, no merged data set information.'] #3.1 3104
        if bhosp_avi:
            #read dates
            df_admdate = pd.read_excel(fadmdate)
            df_disdate = pd.read_excel(fdisdate)
            df_hosp_hn = pd.read_excel(fhhn) 
            #hospital data
            log_2_1 = "<br/>The hospital admission data file (stored in the same folder as the application file) had:" #3.1 3104
            sHospNum = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_records","Parameters"))
            sHospDatefrom = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Minimum_date","Parameters"))
            sHospDateto = ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Maximum_date","Parameters"))
            log_2_2 = indent1(boldstr(sHospNum) + " records with admission dates ranging from " + boldstr(sHospDatefrom)+ " to " + boldstr(sHospDateto))
            #log_2_3 = fn_dateexample(df_admdate,"Admission date",logger)
            #log_2_4 = fn_dateexample(df_disdate,"Discharge date",logger)
            list_log_2_3 = fn_dateexample(df_admdate,"Admission date",logger)
            list_log_2_4 = fn_dateexample(df_disdate,"Discharge date",logger)
            list_log_2_5 = fn_hnexample(df_hosp_hn,"hospital admission data's hospital number",logger)
            list_log_2_5_1 = ""
            try:
                if int(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_HN_with_leadingzero","Parameters")) > 0:
                    list_log_2_5_1 = indent3("hospital number are some with leading zero.")
            except:
                pass
            log_2_6 = indent2("Missing data.")
            log_2_7 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_or_unknown_admission_date","Parameters"))) + " records are missing or unknown format of admission date.")
            log_2_8 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_or_unknown_discharge_date","Parameters"))) + " records are missing or unknown format of discharge date.")
            log_2_7_1 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_wrong_admission_date","Parameters"))) + " records are wrong date or wrong date format of admission date.")
            log_2_8_1 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_wrong_discharge_date","Parameters"))) + " records are wrong date or wrong date format of discharge date")
            log_2_9 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_outcome_result","Parameters"))) + " records are missing outcome.")
            log_2_10 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_age","Parameters"))) + " records are missing age.")
            log_2_11 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_missing_gender","Parameters"))) + " records are missing gender.")
            #log_2 = [log_2_1,log_2_2 ,log_2_3,log_2_4,log_2_5,log_2_6,log_2_7,log_2_8,log_2_9,log_2_10,log_2_11]
            log_2 = [log_2_1,log_2_2] + list_log_2_3 + list_log_2_4 + list_log_2_5 + [list_log_2_5_1] + [log_2_6,log_2_7,log_2_8,log_2_7_1,log_2_8_1,log_2_9,log_2_10,log_2_11]
            #Merge data
            log_3_1 = "<br/>The merged data set between microbiology data and hospital admission data had:"
            log_3_2 = indent2("Microbiology data records unable to merged with hospital admission data records.")
            tempstr =  ""
            try:
                mm = int(ALogL.retrieve_results(over_raw,"microbiology_data","Type_of_data_file","Number_of_HN_with_leadingzero","Parameters"))
                mh = int(ALogL.retrieve_results(over_raw,"hospital_admission_data","Type_of_data_file","Number_of_HN_with_leadingzero","Parameters"))
                if ((mm>0) and (mh<=0)): 
                    tempstr = boldstr("This may due to microbiology data's hospital number are some with leading zero while hospital admission data's hospital number doesn't have.")
                elif ((mm<=0) and (mh>0)):
                    tempstr = boldstr("This may due to hospital admission data's hospital number are some with leading zero while microbiology data's hospital number doesn't have.")
            except:
                pass
            log_3_3 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_unmatchhn","Parameters"))) + " records are unable to merge due to no hospital number found in hospital admission data." + tempstr)
            log_3_4 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_unmatchperiod","Parameters"))) + " records are unable to merge due to have hospital number found in hospital admission data but collection date not in admission period.")
            log_3_5 = indent2("Merged data.")
            log_3_6 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_matchall","Parameters"))) + " records are merged (All specimen type).")
            log_3_7 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_matchblood","Parameters"))) + " records are merged (Only blood specimen).")
            log_3_8 = indent3(boldstr(ALogL.assign_na_toinfo(ALogL.retrieve_results(over_raw,"merged_data","Type_of_data_file","Number_of_matchbsi","Parameters"))) + " records are merged (Only BSI specimen).")
            log_3 = [log_3_1,log_3_2,log_3_3,log_3_4,log_3_5,log_3_6,log_3_7,log_3_8]
        if bhosp_avi:
            log_all = log_1
            ALogL.report_title(c,'Data verification log file',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
            ALogL.report_context(c,log_all, 1.07*inch, 1.2*inch, 460, 650, font_size=11)
            c.showPage()
            log_all = log_2 + log_3
            ALogL.report_context(c,log_all, 1.07*inch, 1.2*inch, 460, 650, font_size=11)
            c.showPage()
        else:
            log_all = log_1 + log_2 + log_3
            ALogL.report_title(c,'Data verification log file',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
            ALogL.report_context(c,log_all, 1.07*inch, 1.2*inch, 460, 650, font_size=11)
            c.showPage()
    except Exception as e:
        AL.printlog("Error : generate data verification log summary page",True,logger)
        logger.exception(e)

    
def tsnodata(c,title_name,title_sub,nodatatext):
    log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
    ALogL.report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
    log_2 = nodatatext
    ALogL.report_context(c,[log_2], 1.0*inch, 8.5*inch, 460, 80, font_size=9)
    c.showPage()
def ts_withfootnote(c,df, df_col, marked_idx, title_name, title_sub,footnote_txt):
    if len(marked_idx) == 1: # no.row < 30 rows
        log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
        df_1 = df.loc[0:]
        df_1 = df_1.values.tolist()
        df_1 = df_col + df_1
        ALogL.report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
        table = df_1
        table_draw = ALogL.report_table_appendix(table)
        table_draw.wrapOn(c, 500, 300)
        h = (30-len(table))*(0.25)
        table_draw.drawOn(c, 1.07*inch, (h+2.1)*inch)
        ALogL.report_context(c,[footnote_txt], 1.0*inch, 0.1*inch, 460, 120, font_size=9,line_space=12)
        c.showPage()
    else:                        # no.row >= 30 rows
        for i in range(len(marked_idx)):
            if i == 0:
                log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:marked_idx[i+1]]
            elif i+1 == len(marked_idx):
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:]
            else:
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]+1:marked_idx[i+1]]
            df_1 = df_1.values.tolist()
            df_1 = df_col + df_1
            ALogL.report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
            table = df_1
            table_draw = ALogL.report_table_appendix(table)
            table_draw.wrapOn(c, 500, 300)
            h = (30-len(table))*(0.25)
            table_draw.drawOn(c, 1.07*inch, (h+2.1)*inch)
            ALogL.report_context(c,[footnote_txt], 1.0*inch, 0.1*inch, 460, 120, font_size=9,line_space=12)
            c.showPage()
def ts(c,df, df_col, marked_idx, title_name, title_sub):
    if len(marked_idx) == 1: # no.row < 30 rows
        log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
        df_1 = df.loc[0:]
        df_1 = df_1.values.tolist()
        df_1 = df_col + df_1
        ALogL.report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
        table = df_1
        table_draw = ALogL.report_table_appendix(table)
        table_draw.wrapOn(c, 500, 300)
        h = (30-len(table))*(0.25)
        table_draw.drawOn(c, 1.07*inch, (h+2.1)*inch)
        c.showPage()
    else:                        # no.row >= 30 rows
        for i in range(len(marked_idx)):
            if i == 0:
                log_1 = "<b>" + title_name + ": " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:marked_idx[i+1]]
            elif i+1 == len(marked_idx):
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]:]
            else:
                log_1 = "<b>" + title_name + " (continue): " + title_sub + "</b>"
                df_1 = df.loc[marked_idx[i]+1:marked_idx[i+1]]
            df_1 = df_1.values.tolist()
            df_1 = df_col + df_1
            ALogL.report_context(c,[log_1], 1.0*inch, 10.0*inch, 460, 80, font_size=11)
            table = df_1
            table_draw = ALogL.report_table_appendix(table)
            table_draw.wrapOn(c, 500, 300)
            h = (30-len(table))*(0.25)
            table_draw.drawOn(c, 1.07*inch, (h+2.1)*inch)
            c.showPage()
"""
logger = logging.getLogger('AMASS_logfile_version_2.py')
logger.setLevel(logging.INFO)
fh = logging.FileHandler("./error_logfile_amass.txt")
fh.setLevel(logging.ERROR)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)
"""

path = AC.CONST_PATH_ROOT
dict_i = path + "dictionary_for_microbiology_data.xlsx"
dict_hosp_i = path + "dictionary_for_hospital_admission_data.xlsx"
dict_ward_i = path + "dictionary_for_wards.xlsx"
over_i = AC.CONST_PATH_RESULT + "logfile_results.csv"
org_i = AC.CONST_PATH_RESULT + "logfile_organism.xlsx"
spc_i = AC.CONST_PATH_RESULT + "logfile_specimen.xlsx"
ast_i = AC.CONST_PATH_RESULT + "logfile_ast.xlsx"
ris_i = AC.CONST_PATH_RESULT + "logfile_ris_count.xlsx"
gen_i = AC.CONST_PATH_RESULT + "logfile_gender.xlsx"
age_i = AC.CONST_PATH_RESULT + "logfile_age.xlsx"
dis_i = AC.CONST_PATH_RESULT + "logfile_discharge.xlsx"
var_mi_i = path + "Variables/variables_micro.xlsx"
var_ho_i = path + "Variables/variables_hosp.xlsx"
var_mi_icsv = path + "Variables/variables_micro.csv"
var_ho_icsv = path + "Variables/variables_hosp.csv"

# file_spc_date = "ResultData/logfile_formatspecdate.xlsx"
# file_adm_date = "ResultData/logfile_formatadmdate.xlsx"
# file_dis_date = "ResultData/logfile_formatdisdate.xlsx"
# file_microhn = "ResultData/logfile_microhn.xlsx"
# file_hosphn = "ResultData/logfile_hosphn.xlsx"
file_spc_date = AC.CONST_PATH_RESULT + "logfile_formatspecdate.xlsx"
file_adm_date = AC.CONST_PATH_RESULT + "logfile_formatadmdate.xlsx"
file_dis_date = AC.CONST_PATH_RESULT + "logfile_formatdisdate.xlsx"
file_microhn = AC.CONST_PATH_RESULT + "logfile_microhn.xlsx"
file_hosphn = AC.CONST_PATH_RESULT + "logfile_hosphn.xlsx"
#file_dup = "ResultData/logfile_dup.xlsx"
file_dup = AC.CONST_PATH_RESULT + "logfile_dup.xlsx"
file_ward = AC.CONST_PATH_RESULT + "logfile_ward.xlsx"
file_ward_hosp = AC.CONST_PATH_RESULT + "logfile_ward_hosp.xlsx"




scol_df_summary_priority = 'priority'
scol_df_summary_valgroup = 'amass_valgroup'
scol_df_summary_count = 'counts'

## Init log 
logger = AL.initlogger('AMR anlaysis',"./log_dataverification_log.txt")
AL.printlog("AMASS version : " + AC.CONST_SOFTWARE_VERSION,False,logger)
AL.printlog("Pandas library version : " + str(pd.__version__),False,logger)
df_summary_amassval = pd.DataFrame(columns=['amass_name',scol_df_summary_count,scol_df_summary_priority,scol_df_summary_valgroup])
try:
    over_raw = pd.read_csv(over_i)
except:
    over_raw = pd.DataFrame(["microbiology_data","microbiology_data","microbiology_data","microbiology_data", 
                             "hospital_admission_data","hospital_admission_data","hospital_admission_data","hospital_admission_data","hospital_admission_data", 
                             "merged_data","merged_data","merged_data","merged_data","merged_data","merged_data"],index=range(15), columns=["Type_of_data_file"])
    over_raw["Parameters"] =["Number_of_missing_specimen_date","Number_of_missing_specimen_type","Number_of_missing_culture_result","format_of_specimen_date",
                             "Number_of_missing_admission_date","Number_of_missing_discharge_type","Number_of_missing_outcome_result","format_of_admission_date","format_of_discharge_date", 
                             "Number_of_missing_specimen_date","Number_of_missing_admission_date","Number_of_missing_discharge_type","Number_of_missing_age","Number_of_missing_gender","Number_of_missing_infection_origin_data",
                             "Number_of_wrong_admission_date","Number_of_wrong_discharge_date","Number_of_wrong_specimen_date"]
    over_raw["Values"] = "NA"

if ALogL.checkpoint(dict_i):
    try:
        s_dict_column =["amass_name","user_name","requirements","explanations"]
        dict_raw = AL.readxlsorcsv_noheader(path,"dictionary_for_microbiology_data",s_dict_column,logger)
        file_format = ALogL.retrieve_uservalue(dict_raw, "file_format")
        lst_opt_2 = ["1", 
                    "hospital_number", ALogL.retrieve_uservalue(dict_raw, "hospital_number"), 
                    "specimen_collection_date", ALogL.retrieve_uservalue(dict_raw, "specimen_collection_date"), 
                    "specimen_type", ALogL.retrieve_uservalue(dict_raw, "specimen_type"), 
                    "organism", ALogL.retrieve_uservalue(dict_raw, "organism"), 
                    "specimen_number", ALogL.retrieve_uservalue(dict_raw, "specimen_number"), 
                    "antibiotic", ALogL.retrieve_uservalue(dict_raw, "antibiotic"), 
                    "ast_result", ALogL.retrieve_uservalue(dict_raw, "ast_result")]
        lst_opt_2 = [x for x in lst_opt_2 if pd.isna(x)==False] #selecting only variables which are not NaN
        #Retrieving all antibiotic list from dictionary_for_microbiology_data.xlsx
        idx_micro_drug = dict_raw.iloc[:,0].tolist().index("Variable names used for \"antibiotics\" in AMASS") #index of part2 header
        idx_micro_spc = dict_raw.iloc[:,0].tolist().index("Data values of variable used for \"specimen_type\" in AMASS") #index of part3 header
        idx_micro_org_core =  dict_raw.iloc[:,0].tolist().index("Data values of variable \"organism\" in AMASS, which are mainly used for the main report")
        idx_micro_opt =  dict_raw.iloc[:,0].tolist().index("Optional: Data values used for the cover of the report generated by the AMASS")
        idx_micro_org_opt =  dict_raw.iloc[:,0].tolist().index("Optional: Data values of variable \"organism\" in AMASS, which are mainly used for the annex")
        idx_micro_opt_2 =  dict_raw.iloc[:,0].tolist().index("Optional: Variables name used in AMASS; this is needed only when data is in long format")

        dict_drug = dict_raw.copy().iloc[idx_micro_drug+1:idx_micro_spc,:2].reset_index().drop(columns=['index']).fillna("").rename(columns={"amass_name":"amass_drug","user_name":"user_drug"})
        dict_spc = dict_raw.copy().iloc[idx_micro_spc+1:idx_micro_org_core,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_core = dict_raw.copy().iloc[idx_micro_org_core+1:idx_micro_opt,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_core["tag"] = "A"
        dict_org_opt = dict_raw.copy().iloc[idx_micro_org_opt+1:idx_micro_opt_2,:2].reset_index().drop(columns=['index']).fillna("")
        dict_org_opt["tag"] = "B"
        dict_org = pd.concat([dict_org_core,dict_org_opt],axis=0)
        #Trim 
        dict_drug["user_drug"] = dict_drug["user_drug"].astype("string").str.strip().fillna("")
        dict_spc["user_name"] = dict_spc["user_name"].astype("string").str.strip().fillna("")
        dict_org["user_name"] = dict_org["user_name"].astype("string").str.strip().fillna("")
        
    except Exception as e:
        logger.exception(e)
        pass

    marked_idx_org = []
    try: #TS1
        try:
            org_raw = pd.read_excel(org_i)
        except:
            org_raw = pd.read_csv(org_i,encoding='windows-1252')
        org_merge = pd.merge(dict_org,org_raw,how="outer",left_on="user_name",right_on="Organism")
        org_merge["Frequency"] = org_merge["Frequency"].fillna(0).astype(int).astype(str)
        org_merge["amass_name"] = org_merge["amass_name"].fillna("zzzz")
        org_merge = org_merge.fillna("zzzz").replace("","zzzz").sort_values(["amass_name"],ascending=True) #all NaN and unknown values >>>> "zzzz" for ordering organism name
        org_A = org_merge.loc[(org_merge["tag"]=="A")&(org_merge["amass_name"]!="zzzz")].sort_values(["amass_name"],ascending=True).replace("zzzz","").reset_index().drop(columns=["index","tag"])
        for idx in org_A.index:
            if org_A.loc[idx,"user_name"] != "" and org_A.loc[idx,"Organism"] == "":
                org_A.at[idx,"Organism"] = org_A.loc[idx,"user_name"]
            else:
                pass
            org_A.at[idx,"Organism"] = ALogL.prepare_unicode(org_A.loc[idx,"Organism"])
        org_A = org_A.drop(columns=["user_name"]).rename(columns={"amass_name":"Data values of variable \"organism\"\nin AMASS, which are mainly used for\nthe main report","Organism":"Data values of variable\nrecorded for \"organism\" in your\nmicrobiology data file", "Frequency":"Number of\nobservations"})
        marked_org_A = ALogL.marked_idx(org_A)
        org_A_col = [list(org_A.columns)]

        org_B = org_merge.loc[(org_merge["tag"]=="B")|(org_merge["amass_name"]=="zzzz"),:].sort_values(["amass_name"],ascending=True).replace("zzzz","").reset_index().drop(columns=["index","tag"])
        for idx in org_B.index:
            if org_B.loc[idx,"user_name"] != "" and org_B.loc[idx,"Organism"] == "":
                org_B.at[idx,"Organism"] = org_B.loc[idx,"user_name"]
            else:
                pass
            org_B.at[idx,"Organism"] = ALogL.prepare_unicode(org_B.loc[idx,"Organism"])
        org_B = org_B.drop(columns=["user_name"]).rename(columns={"amass_name":"Optional: Data values of variable\n\"organism\" in AMASS, which are \nmainly used for the annex", "Organism":"Data values of the variable\nrecorded for \"organism\" in your\nmicrobiology data file", "Frequency":"Number of\nobservations"})
        marked_org_B = ALogL.marked_idx(org_B)
        org_B_col = [list(org_B.columns)]
    except Exception as e:
        logger.exception(e)
        pass

    marked_idx_spc = []
    
    try: #TS2
        try:
            spc_raw = pd.read_excel(spc_i)
        except:
            spc_raw = pd.read_csv(spc_i,encoding='windows-1252')
        spc_merge = pd.merge(dict_spc,spc_raw,how="outer",left_on="user_name",right_on="Specimen")
        spc_merge["Freq"] = spc_merge["Frequency"].fillna(0).astype(int)
        spc_merge["Frequency"] = spc_merge["Frequency"].fillna(0).astype(int).astype(str)
        spc_merge = spc_merge.fillna("zzzz")
        spc_merge = spc_merge.sort_values(["amass_name"],ascending=True).reset_index().drop(columns=["index"])
        spc_merge = spc_merge.replace(regex=["zzzz"],value="")
        for idx in spc_merge.index:
            if spc_merge.loc[idx,"user_name"] != "" and spc_merge.loc[idx,"Specimen"] == "":
                spc_merge.at[idx,"Specimen"] = spc_merge.loc[idx,"user_name"]
            else:
                pass
            spc_merge.at[idx,"Specimen"] = ALogL.prepare_unicode(spc_merge.loc[idx,"Specimen"])
        #temp_df = spc_merge.groupby(["amass_name"]).size().reset_index(name=scol_df_summary_count)
        temp_df = spc_merge.groupby(["amass_name"])["Freq"].sum().reset_index(name=scol_df_summary_count)
        
        
        temp_df["amass_name"] = temp_df["amass_name"].fillna("")
        temp_df[scol_df_summary_priority] = 1
        temp_df.loc[temp_df["amass_name"]=="", scol_df_summary_priority] =999
        temp_df[scol_df_summary_valgroup] = "Specimen type"
        spc_merge.drop(columns=["Freq"],inplace=True)
        df_summary_amassval = pd.concat([df_summary_amassval,temp_df],ignore_index = True)
        spc_merge = spc_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable used\nfor \"specimen type\" in AMASS",
                                                "Specimen":"Data values of variable\nrecorded for \"specimen type\" in your\nmicrobiology data file", 
                                                "Frequency":"Number of\nobservations"})
        marked_spc = ALogL.marked_idx(spc_merge)
        spc_col = [list(spc_merge.columns)]
    except Exception as e:
        logger.exception(e)
        pass
    
    try: #TS3
        try:
            ast_raw = pd.read_excel(ast_i)
        except:
            ast_raw = pd.read_csv(ast_i,encoding='windows-1252')
        #print(dict_drug)
        #print(ast_raw)
        ast_raw.columns = ["user_name", "frequency_raw"]
        ast_merge = pd.merge(dict_drug,ast_raw,how="outer",left_on="user_drug",right_on="user_name")
        ast_merge[["user_drug","user_name"]] = ast_merge[["user_drug","user_name"]].fillna("")
        for idx in ast_merge.index:
            if ast_merge.loc[idx,"amass_drug"] == ast_merge.loc[idx,"user_name"]:
                pass
            else:
                if ast_merge.loc[idx,"user_drug"] != "" and ast_merge.loc[idx,"user_name"] == "":
                    pass
                else:
                    ast_merge.at[idx,"user_drug"] = ast_merge.loc[idx,"user_name"]
        ast_merge["frequency_raw"] = ast_merge["frequency_raw"].fillna(0).astype(int)
        ast_merge["amass_drug"] = ast_merge["amass_drug"].fillna("zzzz")
        ast_merge = ast_merge.sort_values(["amass_drug"],ascending=True).reset_index().drop(columns=["index"])
        ast_merge["amass_drug"] = ast_merge["amass_drug"].replace(regex=["zzzz"],value="")
        ast_merge = ast_merge.loc[~ast_merge["user_drug"].isin(lst_opt_2),:] #excluding variables out
        for idx in ast_merge.index:
            ast_merge.at[idx,"user_drug"] = ALogL.prepare_unicode(ast_merge.loc[idx,"user_drug"])
        #Build 3027 #TS3 unmapped S\I\R values.
        ast_merge["fdiff"] = -1
        ris_raw = pd.DataFrame(columns =["amass_atb", "frequency_raw_ris"])
        try:
            try:
                ris_raw = pd.read_excel(ris_i)
            except:
                ris_raw = pd.read_csv(ris_i,encoding='windows-1252')
            ris_raw.columns = ["amass_atb", "frequency_raw_ris"]
            ast_merge = pd.merge(ast_merge,ris_raw,how="left",left_on="amass_drug",right_on="amass_atb")
            ast_merge["frequency_raw_ris"] = ast_merge["frequency_raw_ris"].fillna(-1).astype(int) 
            ast_merge["fdiff"] = ast_merge["frequency_raw"] - ast_merge["frequency_raw_ris"]
            ast_merge["frequency_raw"] = ast_merge["frequency_raw"].astype(str)
            ast_merge["fdiff"] = ast_merge["fdiff"].astype(str)
            ast_merge.loc[ast_merge["frequency_raw_ris"] == -1,"fdiff"] = "NA"
            ast_merge.loc[ast_merge["amass_drug"].astype(str).str.strip() == "","fdiff"] = ""
            #ast_merge.loc[ast_merge["frequency_raw_ris"] < ast_merge["frequency_raw"], "fdiff"] = ast_merge["frequency_raw"] - ast_merge["frequency_raw_ris"]
        except Exception as e:
            logger.exception(e)
            pass
        try:
            ast_merge = ast_merge.drop(columns=["amass_atb","frequency_raw_ris"])
        except Exception as e:
            logger.exception(e)
            pass
        #ast_merge_for_ris = ast_merge.copy(deep=True)
        ast_merge = ast_merge.reset_index().drop(columns=["index","user_name"]).rename(columns={"amass_drug":"Variable names used for\n\"antibiotics\" described in AMASS",
                                                                        "user_drug":"Variable names recorded for\n\"antibiotics\" in your\nmicrobiology data file",
                                                                        "frequency_raw":"Number of observations\ncontaining S, I, or R\nfor each antibiotic",
                                                                        "fdiff":"Number of observations\ncontaining some data\n(i.e. not blank)\nwhich were not translated\ninto S, I or R*"})
        ast_col = [list(ast_merge.columns)]
        marked_ast = ALogL.marked_idx(ast_merge)
    except Exception as e:
        logger.exception(e)
        pass
##dictionary_for_hospital_admission_data
if ALogL.checkpoint(dict_hosp_i): 
    try:
        try:
            dict_hosp_raw = pd.read_excel(dict_hosp_i).iloc[:,:4]
        except:
            try:
                dict_hosp_raw = pd.read_csv(path + "dictionary_for_hospital_admission_data.csv").iloc[:,:4]
            except:
                dict_hosp_raw = pd.read_csv(path + "dictionary_for_hospital_admission_data.csv",encoding="windows-1252").iloc[:,:4]
        dict_hosp_raw.columns = ["amass_name","user_name","requirements","explanations"]
        
        #Retrieving column names of hospital_admission_data.xlsx 
        male = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="male",["amass_name","user_name"]]
        female = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="female",["amass_name","user_name"]]
        dict_gen = pd.concat([male,female],axis=0)
        died = dict_hosp_raw.loc[dict_hosp_raw["amass_name"]=="died",["amass_name","user_name"]]
        #trim
        dict_gen["user_name"] = dict_gen["user_name"].astype("string").str.strip().fillna("")
        died["user_name"] = died["user_name"].astype("string").str.strip().fillna("")
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS4
        try:
            gen_raw = pd.read_excel(gen_i)
        except:
            gen_raw = pd.read_csv(gen_i,encoding='windows-1252')
        gen_raw["Gender"] = gen_raw["Gender"].astype("string").str.strip().fillna("")
        gen_merge = pd.merge(dict_gen,gen_raw,how="outer",left_on="user_name",right_on="Gender")
        gen_merge["Frequency"] = gen_merge["Frequency"].fillna(0).astype(int).astype(str)
        gen_merge["user_name"] = gen_merge["user_name"].fillna("")
        gen_merge["Gender"] = gen_merge["Gender"].fillna("")
        for idx in gen_merge.index:
            if (gen_merge.loc[idx,"user_name"] != "") and (gen_merge.loc[idx,"Gender"] == ""):
                gen_merge.at[idx,"Gender"] = gen_merge.loc[idx,"user_name"]
            else:
                pass
            gen_merge.at[idx,"Gender"] = ALogL.prepare_unicode(gen_merge.loc[idx,"Gender"])
        gen_merge = gen_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable\n used for \"gender\" described\nin AMASS",
                                                        "Gender":"Data values of variable\n recorded for \"gender\" in your\nhospital admission data file", 
                                                        "Frequency":"Number of\nobservations"})
        gen_col = [list(gen_merge.columns)]
        gen_1 = gen_merge.values.tolist()
        gen_1 = gen_col + gen_1
        marked_gen = ALogL.marked_idx(gen_merge)
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS5
        try:
            age_raw = pd.read_excel(age_i)
            age_raw["Age"] = age_raw["Age"].fillna("NA")
        except:
            age_raw = pd.read_csv(age_i,encoding='windows-1252')
        age_raw["Age_cat"] = ""
        for idx in age_raw.index:
            if age_raw.loc[idx,"Age"] == "NA":
                age_raw.at[idx,"Age_cat"] = "Not available"
            elif int(age_raw.loc[idx,"Age"]) < 1:
                age_raw.at[idx,"Age_cat"] = "Less than 1 year"
            elif int(age_raw.loc[idx,"Age"]) >= 1 and int(age_raw.loc[idx,"Age"]) <= 4:
                age_raw.at[idx,"Age_cat"] = "1 to 4 years"
            elif int(age_raw.loc[idx,"Age"]) >= 5 and int(age_raw.loc[idx,"Age"]) <= 14:
                age_raw.at[idx,"Age_cat"] = "5 to 14 years"
            elif int(age_raw.loc[idx,"Age"]) >= 15 and int(age_raw.loc[idx,"Age"]) <= 24:
                age_raw.at[idx,"Age_cat"] = "15 to 24 years"
            elif int(age_raw.loc[idx,"Age"]) >= 25 and int(age_raw.loc[idx,"Age"]) <= 34:
                age_raw.at[idx,"Age_cat"] = "25 to 34 years"
            elif int(age_raw.loc[idx,"Age"]) >= 35 and int(age_raw.loc[idx,"Age"]) <= 44:
                age_raw.at[idx,"Age_cat"] = "35 to 44 years"
            elif int(age_raw.loc[idx,"Age"]) >= 45 and int(age_raw.loc[idx,"Age"]) <= 54:
                age_raw.at[idx,"Age_cat"] = "45 to 54 years"
            elif int(age_raw.loc[idx,"Age"]) >= 55 and int(age_raw.loc[idx,"Age"]) <= 64:
                age_raw.at[idx,"Age_cat"] = "55 to 64 years"
            elif int(age_raw.loc[idx,"Age"]) >= 65 and int(age_raw.loc[idx,"Age"]) <= 80:
                age_raw.at[idx,"Age_cat"] = "65 to 80 years"
            else:
                age_raw.at[idx,"Age_cat"] = "More than 80 years"
        age_raw["Frequency"] = age_raw["Frequency"].fillna(0).astype(int).astype(str)
        age_merge = age_raw.copy().loc[:,["Age_cat","Age","Frequency"]].fillna("").rename(columns={"Age_cat":"Data values of variable\n used for \"age\" described\nin AMASS",
                                                                                            "Age":"Data values of variable\nrecorded for \"age\" in your\nhospital admission data file", 
                                                                                            "Frequency":"Number of\nobservations"})
        age_col = [list(age_merge.columns)]
        marked_age = ALogL.marked_idx(age_merge)
    except Exception as e:
        logger.exception(e)
        pass

    try: #TS6
        try:
            dis_raw = pd.read_excel(dis_i)
        except:
            dis_raw = pd.read_csv(dis_i,encoding='windows-1252')
        
        dis_raw["Discharge status"] = dis_raw["Discharge status"].astype("string").str.strip().fillna("")
        try:
            dis_raw["Discharge status"] = dis_raw["Discharge status"].str.replace(r'\.0$','',regex=True)
        except:
            pass
        dis_merge = pd.merge(died,dis_raw,how="outer",left_on="user_name",right_on="Discharge status")
        dis_merge["Frequency"] = dis_merge["Frequency"].fillna(0).astype(int).astype(str)
        dis_merge["user_name"] = dis_merge["user_name"].fillna("")
        dis_merge["Discharge status"] = dis_merge["Discharge status"].fillna("")
        for idx in dis_merge.index:
            if (dis_merge.loc[idx,"user_name"] != "") and (dis_merge.loc[idx,"Discharge status"] == ""):
                dis_merge.at[idx,"Discharge status"] = dis_merge.loc[idx,"user_name"]
            else:
                pass
            dis_merge.at[idx,"Discharge status"] = ALogL.prepare_unicode(dis_merge.loc[idx,"Discharge status"])
        temp_df = dis_merge.groupby(["amass_name"]).size().reset_index(name=scol_df_summary_count)
        temp_df[scol_df_summary_priority] = 6
        temp_df[scol_df_summary_valgroup] = "Discharge status"
        df_summary_amassval = pd.concat([df_summary_amassval,temp_df],ignore_index = True)
        dis_merge = dis_merge.drop(columns=["user_name"]).fillna("").rename(columns={"amass_name":"Data values of variable\nname used for \"discharge status\"\ndescribed in AMASS",
                                                        "Discharge status":"Data values of variable\nname recorded for \"discharge status\"\nin your hospital admission data file", 
                                                        "Frequency":"Number of\nobservations"})
        dis_col = [list(dis_merge.columns)]
        marked_dis = ALogL.marked_idx(dis_merge)
    except Exception as e:
        logger.exception(e)
        pass


if (ALogL.checkpoint(var_mi_i) or ALogL.checkpoint(var_mi_icsv)): #TS7
    try:
        try:
            var_mi_raw = pd.read_excel(var_mi_i).rename(columns={"variables_micro":"Variable names used in your microbiology data file"})
        except:
            #var_mi_i = path+"Variables/variables_micro.csv"
            var_mi_raw = pd.read_csv(var_mi_icsv,encoding='utf-8').rename(columns={"variables_micro":"Variable names used in your microbiology data file"})
        var_mi_col = [list(var_mi_raw.columns)]
        marked_var_mi = ALogL.marked_idx(var_mi_raw)
    except Exception as e:
        logger.exception(e) # Will send the errors to the file
        pass

if (ALogL.checkpoint(var_ho_i) or ALogL.checkpoint(var_ho_icsv)): #TS8
    try: #TS8
        try:
            var_ho_raw = pd.read_excel(var_ho_i).rename(columns={"variables_hosp":"Variable names used in your hospital admission data file"})
        except:
            #var_ho_i = path+"Variables/variables_hosp.csv"
            var_ho_raw = pd.read_csv(var_ho_icsv,encoding='utf-8').rename(columns={"variables_hosp":"Variable names used in your hospital admission data file"})
        var_ho_col = [list(var_ho_raw.columns)]
        marked_var_ho = ALogL.marked_idx(var_ho_raw)
    except Exception as e:
        logger.exception(e)
        pass
##Summary tables 
if ALogL.checkpoint(spc_i):
    try:#TS2 sum
        df_spectype_sum = df_summary_amassval[df_summary_amassval[scol_df_summary_valgroup]=="Specimen type"]
        
        df_spectype_sum = df_spectype_sum.sort_values([scol_df_summary_priority],ascending=True).reset_index().drop(columns=["index"])
        df_spectype_sum = df_spectype_sum[['amass_name',scol_df_summary_count]]
        df_spectype_sum = df_spectype_sum.rename(columns={'amass_name':"Data values of variable\nname recorded for \"specimen type\"\nin your hospital admission data file",
                                        scol_df_summary_count:"Number of\nobservations"})
        
        mark_spectype_sum = ALogL.marked_idx(df_spectype_sum )
        
    except Exception as e:
        logger.exception(e)
        pass
if ALogL.checkpoint(file_dup):
    try:#TS9
        df_dup = pd.read_excel(file_dup)    
        df_dup = df_dup.sort_values(["dataval"],ascending=True).reset_index().drop(columns=["index"])
        df_dup = df_dup.rename(columns={"dataval":"Data values of variable name in your data files", "amassvar":"Data values of variable name described in AMASS"})      
        marked_dup = ALogL.marked_idx(df_dup)      
    except Exception as e:
        logger.exception(e)
        pass
#TS10
df_ward = pd.DataFrame()
df_ward_hosp = pd.DataFrame()
col_wardcode = 'Data values of variable\nused for \"ward\" in AMASS'
col_count = "Number of\nobservations"
try:
    #load dictionary ward
    df_dict_ward = pd.DataFrame()
    if AL.checkxlsorcsv(AC.CONST_PATH_ROOT,"dictionary_for_wards"):
        try:
            df_dict_ward = AL.readxlsorcsv_noheader_forceencode(AC.CONST_PATH_ROOT,"dictionary_for_wards", [AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"WARDTYPE","REQ","EXPLAINATION"],"utf-8",logger)
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip() != ""]
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_AMASS].str.startswith("ward_")]
            df_dict_ward = df_dict_ward[[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL]]
            df_dict_ward[AC.CONST_DICTCOL_DATAVAL] =  df_dict_ward[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip().fillna("")
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_AMASS]!= 'ward_hosp']
            if ALogL.checkpoint(file_ward):
                try:#TS10 Micro
                    col_wardname = 'Data values of variable\nrecorded for \"ward\"\nin your microbiology data file'
                    col_wardname_hosp = 'Data values of variable\nrecorded for \"ward\"\nin your hospital data file'
                    df_ward = pd.read_excel(file_ward)  
                    df_ward[AC.CONST_NEWVARNAME_WARDCODE] = df_ward[AC.CONST_NEWVARNAME_WARDCODE].astype("string").str.strip().fillna("")
                    df_ward[AC.CONST_VARNAME_WARD] = df_ward[AC.CONST_VARNAME_WARD].astype("string").str.strip().fillna("")
                    df_ward = df_ward.merge(df_dict_ward, how="outer", left_on=[AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_VARNAME_WARD], right_on=[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL],suffixes=("", "CO"))
                    df_ward[col_wardcode] = df_ward[AC.CONST_DICTCOL_AMASS].fillna(df_ward[AC.CONST_NEWVARNAME_WARDCODE]).fillna("")
                    df_ward[col_wardname] = df_ward[AC.CONST_DICTCOL_DATAVAL].fillna(df_ward[AC.CONST_VARNAME_WARD])
                    df_ward[col_count] = df_ward["Count"].fillna(0).astype(int).astype(str)
                    df_ward["gotcode"] = 0 
                    df_ward.loc[df_ward[col_wardcode] =="", "gotcode"] = 1
                    df_ward = df_ward.sort_values(["gotcode",col_wardcode],ascending=True).reset_index().drop(columns=["index"])
                    df_ward = df_ward[[col_wardcode,col_wardname,col_count]]
                    AL.fn_savexlsx(df_ward, AC.CONST_PATH_RESULT + "logfile_TS10A_ward.xlsx", logger)
                    marked_ward = ALogL.marked_idx(df_ward)  
                    print("Generated TS10A")
                except Exception as e:
                    logger.exception(e)
                    print(e)
                    print("Warning : error generate ward log for microbiology data")
                    pass
            if ALogL.checkpoint(file_ward_hosp):
                try:#TS10 hosp
                    col_wardname = 'Data values of variable\nrecorded for \"ward\"\nin your hospital data file'
                    df_ward_hosp = pd.read_excel(file_ward_hosp)  
                    df_ward_hosp[AC.CONST_NEWVARNAME_WARDCODE_HOSP] = df_ward_hosp[AC.CONST_NEWVARNAME_WARDCODE_HOSP].astype("string").str.strip().fillna("")
                    df_ward_hosp[AC.CONST_VARNAME_WARD_HOSP] = df_ward_hosp[AC.CONST_VARNAME_WARD_HOSP].astype("string").str.strip().fillna("")
                    df_ward_hosp = df_ward_hosp.merge(df_dict_ward, how="outer", left_on=[AC.CONST_NEWVARNAME_WARDCODE_HOSP,AC.CONST_VARNAME_WARD_HOSP], right_on=[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL],suffixes=("", "CO"))
                    df_ward_hosp[col_wardcode] = df_ward_hosp[AC.CONST_DICTCOL_AMASS].fillna(df_ward_hosp[AC.CONST_NEWVARNAME_WARDCODE_HOSP]).fillna("")
                    df_ward_hosp[col_wardname] = df_ward_hosp[AC.CONST_DICTCOL_DATAVAL].fillna(df_ward_hosp[AC.CONST_VARNAME_WARD_HOSP])
                    df_ward_hosp[col_count] = df_ward_hosp["Count"].fillna(0).astype(int).astype(str)
                    df_ward_hosp["gotcode"] = 0 
                    df_ward_hosp.loc[df_ward_hosp[col_wardcode] =="", "gotcode"] = 1
                    df_ward_hosp = df_ward_hosp.sort_values(["gotcode",col_wardcode],ascending=True).reset_index().drop(columns=["index"])
                    df_ward_hosp = df_ward_hosp[[col_wardcode,col_wardname,col_count]]
                    AL.fn_savexlsx(df_ward_hosp, AC.CONST_PATH_RESULT + "logfile_TS10B_ward_hosp.xlsx", logger)
                    marked_ward_hosp = ALogL.marked_idx(df_ward_hosp)  
                    print("Generated TS10B")
                except Exception as e:
                    logger.exception(e)
                    print(e)
                    print("Warning : error generate ward log for hospital data")
                    pass
        except Exception as e:
            logger.exception(e)
            print("Warning : error generate ward log")
            pass
except Exception as e:
    logger.exception(e)
    print("Warning : error generate ward log")
    pass
##Generating logfile_amass.pdf and Exporting logfile.xlsx
try:
    c = canvas.Canvas(path + "Data_verification_logfile_report.pdf")
    bhosp_avi  = ALogL.checkxlsorcsv(path,"hospital_admission_data")
    if ALogL.checkpoint(over_i):
        try:
            cover(c,over_raw)
            ts0(c,logger,over_raw,file_spc_date,file_adm_date,file_dis_date,file_microhn,file_hosphn,bhosp_avi)
        except Exception as e:
            logger.exception(e)
            pass
    if ALogL.checkpoint(dict_i):
        if ALogL.checkpoint(org_i):
            try:
                ts(c,org_A, [list(org_A.columns)], marked_org_A, "Table S1A", "List of data values of the variable recorded for \"organism\" in your microbiology data file, which are mainly used for the main report")
                org_A.columns = [w.replace("\n"," ") for w in org_A.columns.tolist()]
                org_A.to_excel(AC.CONST_PATH_RESULT + "logfile_TS1A_main_organisms.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
            try:
                ts(c,org_B, [list(org_B.columns)], marked_org_B, "Table S1B", "List of data values of the variable recorded for \"organism\" in your microbiology data file, which are mainly used for the annex")
                org_B.columns = [w.replace("\n"," ") for w in org_B.columns.tolist()]
                org_B.to_excel(AC.CONST_PATH_RESULT + "logfile_TS1B_optional_organisms.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if ALogL.checkpoint(spc_i):
            try:
                ts(c,df_spectype_sum, [list(df_spectype_sum.columns)], mark_spectype_sum, "Table S2 (Summary)", "List of number of records per AMASS's \"specimen type\" in your microbiology data file") #3.1 3104
                df_spectype_sum.columns = [w.replace("\n"," ") for w in df_spectype_sum.columns.tolist()]
                df_spectype_sum.to_excel(AC.CONST_PATH_RESULT + "logfile_TS2SUM_specimens.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
            try:
                ts(c,spc_merge, [list(spc_merge.columns)], marked_spc, "Table S2", "List of data values of the variable recorded for \"specimen type\" in your microbiology data file") #3.1 3104
                spc_merge.columns = [w.replace("\n"," ") for w in spc_merge.columns.tolist()]
                spc_merge.to_excel(AC.CONST_PATH_RESULT + "logfile_TS2_specimens.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
            
        if ALogL.checkpoint(ast_i):
            try:
                ts_withfootnote(c,ast_merge, [list(ast_merge.columns)], marked_ast, "Table S3", "List of variables recorded for \"antibiotics\" in your microbiology data file",
                                "* For antibiotics used for the AST, the numbers in this column should be 0. This is because there should be no data that is not translated into S, I or R. This may occur when laboratories record some number into the same column (e.g. a mix of S I R and zone size) or when dictionary is still incomplete (e.g. recording “R” and “R - no zone” but having only “R” in the dictionary); NA=Not applicable. This could occur when the variable names for antibiotics are not used in AMASS.")
                ast_merge.columns = [w.replace("\n"," ") for w in ast_merge.columns.tolist()]
                ast_merge.to_excel(AC.CONST_PATH_RESULT + "logfile_TS3_antibiotics.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
    if ALogL.checkpoint(dict_hosp_i):
        if ALogL.checkpoint(gen_i):
            try:
                ts(c,gen_merge, [list(gen_merge.columns)], marked_gen, "Table S4", "List of data values of variable recorded for \"gender\" in your hospital admission data file")
                gen_merge.columns = [w.replace("\n"," ") for w in gen_merge.columns.tolist()]
                gen_merge.to_excel(AC.CONST_PATH_RESULT + "logfile_TS4_gender.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        if (ALogL.checkpoint(age_i) and (len(age_merge) > 0)):
            try:
                ts(c,age_merge, [list(age_merge.columns)], marked_age, "Table S5", "List of data values of variable recorded for \"age\" in your hospital admission data file")
                age_merge.columns = [w.replace("\n"," ") for w in age_merge.columns.tolist()]
                age_merge.to_excel(AC.CONST_PATH_RESULT + "logfile_TS5_age.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
        else:
            try:
                tsnodata(c,"Table S5", "List of data values of variable recorded for \"age\" in your hospital admission data file","No age or birth date column in hospital admission data file defined in dictionary or the column defined contains blank data.")
            except Exception as e:
                logger.exception(e)
                pass
        if ALogL.checkpoint(dis_i):
            try:
                ts(c,dis_merge, [list(dis_merge.columns)], marked_dis, "Table S6", "List of data values of variable recorded for \"discharge status\" in your hospital admission data file")
                dis_merge.columns = [w.replace("\n"," ") for w in dis_merge.columns.tolist()]
                dis_merge.to_excel(AC.CONST_PATH_RESULT + "logfile_TS6_discharge_status.xlsx",encoding="UTF-8",index=False,header=True)
            except Exception as e:
                logger.exception(e)
                pass
    if (ALogL.checkpoint(var_mi_i) or ALogL.checkpoint(var_mi_icsv)):
        try:
            ts(c,var_mi_raw, var_mi_col, marked_var_mi, "Table S7", "List of variable names in your microbiology data file") #3.1 3104
        except Exception as e:
            logger.exception(e) 
            pass
    if (ALogL.checkpoint(var_ho_i) or ALogL.checkpoint(var_ho_icsv)):
        try:
            ts(c,var_ho_raw, var_ho_col, marked_var_ho, "Table S8", "List of variable names in your hospital admission data file") #3.1 3104
        except Exception as e:
            logger.exception(e)
            pass
    bisnodup = True
    if ALogL.checkpoint(file_dup):
        try:
            if len(df_dup) > 0:
                ts(c,df_dup, [list(df_dup.columns)], marked_dup, "Table S9", "Duplicate mapping of data value of variable in your with data values of variable describe in AMASS")
                bisnodup = False
        except Exception as e:
            logger.exception(e)
        pass
    #else:
    if  bisnodup:
        try:
            tsnodata(c,"Table S9", "Duplicate mapping of data value of variable in your with data values of variable describe in AMASS","There are no duplicate records in dictionary files that need your revision.")
        except Exception as e:
            logger.exception(e)
            pass
    if len(df_ward) > 0:
        try:
            ts(c,df_ward, [list(df_ward.columns)], marked_ward, "Table S10A", "List of data values of variable recorded for \"ward\" in your microbiology data file")
        except Exception as e:
            logger.exception(e)
            pass
    else:
        try:
            tsnodata(c,"Table S10A", "List of data values of variable recorded for \"ward\" in your microbiology data file","No ward column in microbiology data file defined in dictionary for wards.")
        except Exception as e:
            logger.exception(e)
            pass
    if len(df_ward_hosp) > 0:
        try:
            ts(c,df_ward_hosp, [list(df_ward_hosp.columns)], marked_ward_hosp, "Table S10B", "List of data values of variable recorded for \"ward\" in your hospital admission data file.\n(Data values will be used if missing ward data in microbiology data file for that merged data record.)")
        except Exception as e:
            logger.exception(e)
            pass
    else:
        try:
            tsnodata(c,"Table S10B", "List of data values of variable recorded for \"ward\" in your hospital admission data file.\n(Data values will be used if missing ward data in microbiology data file for that merged data record.)","No ward column in hospital admission data file defined in dictionary for wards.")
        except Exception as e:
            logger.exception(e)
            pass
    c.save()
except Exception as e:
    logger.exception(e)
    pass