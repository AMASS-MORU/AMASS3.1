#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Supplementary data indicators reports systematically.

# Created on 20th April 2022
import logging #for creating error_log
import pandas as pd #for creating and manipulating dataframe
from datetime import date #for generating today date
import datetime  
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_LEFT, TA_CENTER #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
from AMASS_supplementary_report_commonlib import * #for importing data indicators functions

def generate_supplementary_report(df_dict_micro,logger,bisusingmappeddata):
    AL.printlog("Start Supplementary data indicators report: " + str(datetime.datetime.now()),False,logger)
    path = AC.CONST_PATH_ROOT
    summary_res_i = AC.CONST_PATH_RESULT +"Supplementary_data_indicators_results.csv"
    summary_1_i  =AC.CONST_PATH_RESULT +"Supplementary_data_indicators_indicator1.xlsx"
    summary_2_i = AC.CONST_PATH_RESULT +"Supplementary_data_indicators_indicator2.xlsx"
    summary_3a_i = AC.CONST_PATH_RESULT +"Supplementary_data_indicators_indicator3a.xlsx"
    summary_3b_i = AC.CONST_PATH_RESULT +"Supplementary_data_indicators_indicator3b.xlsx"
    summary_list_B_i = AC.CONST_PATH_REPORTWITH_PID +"Report_with_patient_identifiers_annexB_withstatus.xlsx"
    summary_list_A_i = AC.CONST_PATH_REPORTWITH_PID +"Report_with_patient_identifiers_annexA_withstatus.xlsx"
    
    today = date.today()
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=8,alignment=TA_LEFT)
    
    title_summary_1 = "Summary of potential contaminants"
    title_summary_2 = "Summary of notifiable antibiotic-pathogen combinations"
    title_summary_3 = "Summary of infrequent phenotypes or potential errors in AST results based on the indicators that the organisms are intrinsically resistant to an antibiotic but are reported as susceptible"
    title_summary_4 = "Summary of infrequent phenotypes or potential errors in AST results based on the indicators that the isolates exhibit discordant AST results"
    
    note_summary_1 = "Blood culture contamination rate is defined as the number of raw contaminated cultures per number of blood cultures received by the laboratory per reporting period. Blood culture contamination rate will not be estimated in case that the data of negative culture (specified as 'no growth' in the dictionary_for_microbiology_data file) is not available. Details of the criteria are available in \"list_of_indicators.xlsx\" in the folder \"Configuration\"."
    note_summary_2_1 = "Notifiable antibiotic-pathogen combinations and their classifications are defined as WHO list of AMR priority pathogen published in 2017 [1]. The proportion represents the number of patients with blood culture positive for non-susceptible isolates (numerator) over the total number of patient with blood culture positive and AST result available in the raw microbiology data (denominator). Details of the criteria are available in \"list_of_indicators.xlsx\" in the folder \"Configuration\". NS=Non-susceptible; 3GC-NS=3rd-generation cephalosporin; Carbapenems-NS: imipenem, meropenem, ertapenem or doripenem;  Fluoroquinolones-NS: ciprofloxacin or levofloxacin; Methicillin: methicillin, oxacillin, or cefoxitin"
    note_summary_2_2 = "[1] World Health Organization. Global priority list of antibiotic-resistant bacteria to guide research discover, and development of new antibiotics. 2017. https://www.who.int/medicines/publications/WHO-PPL-Short_Summary_25Feb-ET_NM_WHO.pdf. accessed 7th December 2021."
    note_summary_3_1 = "A summary on isolates with infrequent phenotypes that is rarely seen and may potentially be errors in antimicrobial resistant testing results. The proportion represents the number of patients with discordant AST results (numerator) over the total number of patients with blood culture positive and AST result available in the raw microbiology data (denominator). Details of the criteria are available in \"list_of_indicators.xlsx\" in the folder \"Configuration\". AST: antimicrobial-susceptibility test"
    note_summary_3_2 = "*The numerator counts the number of isolates that exhibit discordant AST results between penicillin and beta-lactam combinations. For example, an isolate which is reported as susceptible to amoxicillin but non-susceptible to amoxicillin/clavulanic acid. "
    note_summary_3_3 = "**The numerator counts the number of isolates that exhibit discordant AST results in penicillin antibiotics. For example, an isolate which is reported as is susceptible to ampicillin/sulbactam but non-susceptible to piperacillin/tazobactam OR ticarcillin/clavulanic acid. "
    note_summary_3_4 = "***The numerator counts the number of isolates that exhibit discordant AST results between quinolone and fluoroquinolone. For example, an isolate which is reported as susceptible to nalidixic acid but non-susceptible to fluoroquinolones. "
    note_summary_3_5 = "****The numerator counts the number of Enterobacteriaceae or <i>P. aeruginosa</i> isolates that exhibit discordant AST in aminoglycosides. For example, an Enterobacteriaceae isolate which is reported as non-susceptible to amikacin but susceptible to gentamicin, netilmicin, or tobramycin. "
    note_summary_3_6 = "*****The numerator counts the number of Enterobacteriaceae isolates that exhibit discordant AST in cephems. For example, an Enterobacteriaceae isolate which is reported as susceptible to first generation cephalosporin or second-generation cephalosporin, but non-susceptible to third-generation cephalosporin. "
    
    index_mark_page_bloodcon = []
    index_mark_page_pathogen = []
    index_mark_page_3a = []
    index_mark_page_3b = []
    list_alerted_bloodcon_col = []
    list_alerted_pathogen_col = []
    list_alerted_3a_col = []
    list_alerted_3b_col = []
    list_alerted_bloodcon = []
    list_alerted_pathogen = []
    list_alerted_3a = []
    list_alerted_3b = []
    list_alerted_A = []
    index_mark_page_A = []
    list_alerted_A_col = []
    
    try:
        try:
            config = pd.read_excel(path + "Configuration/Configuration.xlsx")
        except:
            try:
                config = pd.read_csv(path + "Configuration/Configuration.csv")
            except:
                config = pd.read_csv(path + "Configuration/Configuration.csv", encoding="windows-1252")
        if check_config(config, "data_indicators_report"):
            #Reading raw dictionary_for_microbiology_data.xlsx and retrieving user values
            """try:
                dict_micro = pd.read_excel(path + "dictionary_for_microbiology_data.xlsx").iloc[:,:2].fillna("")
            except:
                try:
                    dict_micro = pd.read_csv(path + "dictionary_for_microbiology_data.csv").iloc[:,:2].fillna("")
                except:
                    dict_micro = pd.read_csv(path + "dictionary_for_microbiology_data.csv",encoding="windows-1252").iloc[:,:2].fillna("")
            dict_micro.columns = ["amass_name","user_name"]"""
            #dict_micro = df_dict_micro.iloc[:,:2].fillna("")
            if bisusingmappeddata == True:
                dict_micro = df_dict_micro.iloc[:,:2].fillna("")
            else:            
                print("Supplementary reload dict")
                dict_micro = AL.readxlsxorcsv(path, "dictionary_for_microbiology_data",logger).iloc[:,:2].fillna("")
            dict_micro.columns = ["amass_name","user_name"]
            #The following change because annex b file generated are from microfile after match with dictionary thusm column name changed
            """hn      = retrieve_uservalue(dict_micro, "hospital_number")
            spcdate = retrieve_uservalue(dict_micro, "specimen_collection_date")
            spctype = retrieve_uservalue(dict_micro, "specimen_type")
            spcnum  = retrieve_uservalue(dict_micro, "specimen_number")
            organism= retrieve_uservalue(dict_micro, "organism")"""
            hn      = AC.CONST_VARNAME_HOSPITALNUMBER
            spcdate = AC.CONST_VARNAME_SPECDATERAW
            spctype = AC.CONST_VARNAME_SPECTYPE
            spcnum  = AC.CONST_VARNAME_SPECNUM
            organism= AC.CONST_VARNAME_ORG
            if bisusingmappeddata != True:
                hn      = retrieve_uservalue(dict_micro, "hospital_number")
                spcdate = retrieve_uservalue(dict_micro, "specimen_collection_date")
                spctype = retrieve_uservalue(dict_micro, "specimen_type")
                spcnum  = retrieve_uservalue(dict_micro, "specimen_number")
                organism= retrieve_uservalue(dict_micro, "organism")
            lst_no_growth   = retrieve_userlist(dict_micro,"organism_no_growth")
            font_style = ParagraphStyle('normal',fontName='Helvetica',fontSize=7,alignment=TA_CENTER)
    
            try:
                summary_1 = pd.read_excel(summary_1_i)
                summary_1["except_organism"]  = summary_1["except_organism"].fillna("")
                summary_1["include_organism"] = summary_1["include_organism"].fillna("")
                summary_1["blood_samples"] = summary_1["blood_samples"].fillna("NA")
                summary_1 = summary_1.reset_index().drop(columns=["index"])
                summary_1["organism_fmt"] = ""
                for idx in summary_1.index:
                    summary_1.at[idx,"rule_organism"] = prepare_org_core_v2(summary_1.loc[idx,"rule_organism"].capitalize()).replace("<i>Viridans</i> <i>Group</i> <i>streptococci</i>","Viridans group streptococci")
                    ##add except organism into organism column
                    if summary_1.loc[idx,"except_organism"] != "":
                        summary_1.loc[idx,"except_organism"] = prepare_except_org_for_summarytable(summary_1.loc[idx,"except_organism"])
                        summary_1.at[idx,"organism_fmt"] = Paragraph(summary_1.loc[idx,"rule_organism"] + " except " + summary_1.loc[idx,"except_organism"], style_summary)
                    else:
                        if summary_1.loc[idx,"include_organism"] != "":
                            summary_1.loc[idx,"include_organism"] = prepare_except_org_for_summarytable(summary_1.loc[idx,"include_organism"])
                            summary_1.at[idx,"organism_fmt"] = Paragraph(summary_1.loc[idx,"rule_organism"] + " include " + summary_1.loc[idx,"include_organism"], style_summary)
                        else:
                            summary_1.at[idx,"organism_fmt"] = Paragraph(summary_1.loc[idx,"rule_organism"], style_summary)
                    ##estimated number will not show if no_growth is not available
                    if len(lst_no_growth) > 0:
                        pass
                    else:
                        summary_1.at[:,"blood_samples"] = "NA"
                summary_1 = summary_1.loc[:,["organism_fmt","blood_samples"]]
                summary_1 = summary_1.rename(columns={"organism_fmt":"Organisms","blood_samples":"Proportion of blood samples\n(n)"})
                summary_1_col = [list(summary_1.columns)]
                index_mark_page_summary_1 = create_lst_marked_page(summary_1, 30)
            except Exception as e:
                logger.exception(e)
            pass
    
            try:
                summary_2 = pd.read_excel(summary_2_i).fillna("NA")
                summary_2 = prepare_summarytable(summary_2).rename(columns={"rule_organism":"Organisms","antibiotic":"Antimicrobial-susceptible\nprofile","blood_samples":"Proportion of blood samples\n(n)"})
                summary_2_col = [list(summary_2.columns)]
                index_mark_page_summary_2 = create_lst_marked_page(summary_2, 30)
            except Exception as e:
                logger.exception(e)
                pass
    
            try:
                summary_3a = pd.read_excel(summary_3a_i).fillna("NA")
                summary_3a = prepare_summarytable(summary_3a).rename(columns={"rule_organism":"Organisms","antibiotic":"Antibiotic that intrinsically\nresistant but reported as susceptible","blood_samples":"Proportion of blood samples\n(n)"})
                summary_3a_col = [list(summary_3a.columns)]
                index_mark_page_summary_3a = create_lst_marked_page(summary_3a, 30)
            except Exception as e:
                logger.exception(e)
                pass
    
            try:
                summary_3b = pd.read_excel(summary_3b_i).fillna("NA")
                summary_3b["antibiotic"] = summary_3b["antibiotic"].replace("Penicillins, Betalactam combinations", "Penicillins, Betalactam combinations*").replace("Penicillins", "Penicillins**").replace("Quinolones, Fluoroquinolones", "Quinolones, Fluoroquinolones***").replace("Aminoglycosides", "Aminoglycosides****").replace("Cephems", "Cephems*****")
                summary_3b = prepare_summarytable(summary_3b).rename(columns={"rule_organism":"Organisms","antibiotic":"Antibiotic class that the isolates\nexhibit discordant AST results","blood_samples":"Proportion of blood samples\n(n)"})
                summary_3b_col = [list(summary_3b.columns)]
                index_mark_page_summary_3b = create_lst_marked_page(summary_3b, 30)
            except Exception as e:
                logger.exception(e)
                pass
    
            #####List alerted obs: ANNEX B#####
            if checkpoint(summary_list_B_i):
                list_alerted = pd.read_excel(summary_list_B_i).fillna("")
                list_alerted = format_date_forexportation(df = list_alerted,col_date = spcdate)
                #####Blood Contamination#####
                try:
                    list_alerted_bloodcon = prepare_list_alerted_selectrow(list_alerted, "Possible contaminant", "status_indicator_1")
                    list_alerted_bloodcon = prepare_list_alerted_formatorg(list_alerted_bloodcon)
                    list_alerted_bloodcon = prepare_list_alerted_selectcol(list_alerted_bloodcon, "Possible contaminant",col_hn=hn, col_spcdate="spcdate_fmt")
                    list_alerted_bloodcon_col = [list(list_alerted_bloodcon.columns)]
                    index_mark_page_bloodcon = create_lst_marked_page(list_alerted_bloodcon, 30)
                except Exception as e:
                    logger.exception(e)
                    pass
                #####Notifiable antibiotic-pathogen combination#####
                try:
                    list_alerted_pathogen = list_alerted.loc[(list_alerted["Notifiable antibiotic-pathogen combination"]!="") & (list_alerted["status_indicator_2"]=="yes"),:].fillna("")
                    list_alerted_pathogen["Notifiable antibiotic-pathogen combination"] = list_alerted_pathogen["Notifiable antibiotic-pathogen combination"].replace(regex=[" isolate "],value=" isolate\n").replace(regex=["markednewline\("],value="\n(")
                    list_alerted_pathogen = prepare_list_alerted_selectrow(list_alerted, "Notifiable antibiotic-pathogen combination", "status_indicator_2")
                    list_alerted_pathogen = prepare_list_alerted_formatorg(list_alerted_pathogen,col_oth="Notifiable antibiotic-pathogen combination")
                    list_alerted_pathogen = prepare_list_alerted_selectcol(list_alerted_pathogen, "Notifiable antibiotic-pathogen combination", col_hn=hn, col_spcdate="spcdate_fmt", col_oth="Notifiable antibiotic-pathogen combination",col_refmt_oth="Notifiable antibiotic-pathogen combination")
                    list_alerted_pathogen_col = [list(list_alerted_pathogen.columns)]
                    index_mark_page_pathogen = create_lst_marked_page(list_alerted_pathogen, 10)
                except Exception as e:
                    logger.exception(e)
                    pass
                #####Error in AST result#####
                #3a
                try:
                    list_alerted_3a = list_alerted.copy().rename(columns={"Potential error in either identification or AST result: the species identified usually exhibits intrinsic resistant to the antibiotic but AST result suggested susceptible":"3a"})
                    list_alerted_3a = list_alerted_3a.loc[(list_alerted_3a["3a"]!="") & (list_alerted_3a["status_indicator_3a"]=="yes"),:].fillna("")
                    list_alerted_3a["3a"] = list_alerted_3a["3a"].replace(regex=[" isolate "],value=" isolate\n").replace(regex=["markednewline\("],value="\n(")
                    list_alerted_3a = prepare_list_alerted_selectrow(list_alerted_3a, "3a", "status_indicator_3a")
                    list_alerted_3a = prepare_list_alerted_formatorg(list_alerted_3a,col_oth="3a")
                    list_alerted_3a = prepare_list_alerted_selectcol(list_alerted_3a, "3a", col_hn=hn, col_spcdate="spcdate_fmt", col_oth="3a",col_refmt_oth="Potential errors in the AST results")
                    list_alerted_3a_col = [list(list_alerted_3a.columns)]
                    index_mark_page_3a = create_lst_marked_page(list_alerted_3a, 10)
                except Exception as e:
                    logger.exception(e)
                    pass
                #3b
                try:
                    list_alerted_3b = list_alerted.copy().rename(columns={"Discordant AST results":"3b"})
                    list_alerted_3b = list_alerted_3b.loc[(list_alerted_3b["3b"]!="") & (list_alerted_3b["status_indicator_3b"]=="yes"),:].fillna("")
                    list_alerted_3b["3b"] = list_alerted_3b["3b"].replace(regex=[" isolate "],value=" isolate\n").replace(regex=["markednewline\("],value="\n(")
                    list_alerted_3b = prepare_list_alerted_selectrow(list_alerted_3b, "3b", "status_indicator_3b")
                    list_alerted_3b = prepare_list_alerted_formatorg(list_alerted_3b,col_oth="3b")
                    list_alerted_3b = prepare_list_alerted_selectcol(list_alerted_3b, "3b", col_hn=hn, col_spcdate="spcdate_fmt", col_oth="3b",col_refmt_oth="Potential errors in the AST results")
                    list_alerted_3b_col = [list(list_alerted_3b.columns)]
                    index_mark_page_3b = create_lst_marked_page(list_alerted_3b, 10)
                except Exception as e:
                    logger.exception(e)
                    pass
            else:
                pass
    
            #####List alerted obs: ANNEX A#####
            try:
                list_alerted_A = pd.read_excel(summary_list_A_i).loc[:,[hn, spcdate, "mapped_spctype","mapped_sci"]].fillna("")
                list_alerted_A_col = [list(list_alerted_A.columns)]
                for idx in list_alerted_A.index:
                    list_alerted_A.at[idx,"mapped_sci"] = Paragraph(prepare_org_annexA(list_alerted_A.loc[idx,"mapped_sci"].capitalize()).replace("<i>paratyphi</i>","Paratyphi").replace("<i>typhi</i>","Typhi"),font_style)
                list_alerted_A = format_date_forexportation(df = list_alerted_A,col_date = spcdate)
                spcdate_fmt = "spcdate_fmt"
                list_alerted_A = list_alerted_A.drop(columns=[spcdate]).loc[:,[hn,spcdate_fmt,"mapped_spctype","mapped_sci"]].rename(columns={hn:"Hospital number",spcdate_fmt:"Specimen collection date", "mapped_spctype":"Specimen type", "mapped_sci":"Organisms"})
                list_alerted_A_col = [list(list_alerted_A.columns)]
                index_mark_page_A = create_lst_marked_page(list_alerted_A, 30)
            except Exception as e:
                logger.exception(e)
                list_alerted_A = pd.DataFrame()
    
            try:
                summary = pd.read_csv(summary_res_i)
                spc_date_start_cov = assign_na_toinfo(retrieve_results(summary,"Minimum_date","Parameters"),coverpage=True)
                spc_date_end_cov   = assign_na_toinfo(retrieve_results(summary,"Maximum_date","Parameters"),coverpage=True)
                hospital_name  = assign_na_toinfo(retrieve_results(summary,"Hospital_name","Parameters"),coverpage=True)
                country_name   = assign_na_toinfo(retrieve_results(summary,"Country","Parameters"),coverpage=True)
                spc_date_start = assign_na_toinfo(retrieve_results(summary,"Minimum_date","Parameters"))
                spc_date_end   = assign_na_toinfo(retrieve_results(summary,"Maximum_date","Parameters"))
                num_all_spc    = assign_na_toinfo(retrieve_results(summary,"Number_of_records","Parameters"))
                num_all_spc_pos= assign_na_toinfo(retrieve_results(summary,"Number_of_all_cultue_positive","Parameters"))
                num_blo_spc    = assign_na_toinfo(retrieve_results(summary,"Number_of_blood_specimens_collected","Parameters"))
                num_blo_spc_pos= assign_na_toinfo(retrieve_results(summary,"Number_of_blood_culture_positive","Parameters"))
                num_blo_spc_neg= assign_na_toinfo(retrieve_results(summary,"Number_of_blood_culture_negative","Parameters"))
            except:
                spc_date_start_cov = "Not available"
                spc_date_end_cov   = "Not available"
                hospital_name      = "Not available"
                country_name       = "Not available"
                spc_date_start  = "NA"
                spc_date_end    = "NA"
                num_all_spc     = "NA"
                num_all_spc_pos = "NA"
                num_blo_spc     = "NA"
                num_blo_spc_pos = "NA"
                num_blo_spc_neg = "NA"
        else:
            spc_date_start_cov = "Not available"
            spc_date_end_cov   = "Not available"
            hospital_name      = "Not available"
            country_name       = "Not available"
            spc_date_start  = "NA"
            spc_date_end    = "NA"
            num_all_spc     = "NA"
            num_all_spc_pos = "NA"
            num_blo_spc     = "NA"
            num_blo_spc_pos = "NA"
            num_blo_spc_neg = "NA"
    except Exception as e:
        logger.exception(e)
        pass
    
    def summary_table_page1(page,spc_date_start=spc_date_start,spc_date_end=spc_date_end,
                      num_all_spc=num_all_spc,num_all_spc_pos=num_all_spc_pos, 
                      num_blo_spc=num_blo_spc, num_blo_spc_pos=num_blo_spc_pos, 
                      num_blo_spc_neg=num_blo_spc_neg,today=today.strftime("%d %b %Y")):
        ##paragraph variable
        iden1_op = "<para leftindent=\"35\">"
        iden2_op = "<para leftindent=\"70\">"
        iden_ed = "</para>"
        bold_blue_op = "<b><font color=\"#000080\">"
        bold_blue_ed = "</font></b>"
        add_blankline = "<br/>"
        ##content
        summary_1_1 = "The tables are counts of records of blood samples that violated the data validation indicators stratified by the level of priority as indicated in the list_of_indicators.xlsx."
        summary_1_2 = "In brief, the microbiology data is de-duplicated by including only the first isolate per unique specimen number per specimen type per organism identified per evaluation period."
        summary_1 = [summary_1_1, add_blankline + summary_1_2]
        summary_2_1 = "The microbiology_data file had:"
        summary_2_2 = "Sample collection dates ranged from " + bold_blue_op + str(spc_date_start) + " to " + str(spc_date_end) + bold_blue_ed
        summary_2_3 = "Number of records of all specimen types collected within the above date range:"
        summary_2_4 = bold_blue_op + str(num_all_spc) + " records" + bold_blue_ed
        summary_2_5 = "Number of records of all specimen types with culture positive for a microorganism:"
        summary_2_6 = bold_blue_op + str(num_all_spc_pos) + " records" + bold_blue_ed
        summary_2_7 = "Number of records of blood specimens collected within the above date range:"
        summary_2_8 = bold_blue_op + str(num_blo_spc) + " records" + bold_blue_ed
        summary_2_9 = "Number of records of blood specimens with culture positive for a microorganism:"
        summary_2_10 = bold_blue_op + str(num_blo_spc_pos) + " records" + bold_blue_ed
        summary_2_11 = "Number of records of blood specimens with no growth for a microorganism:"
        summary_2_12 = bold_blue_op + str(num_blo_spc_neg) + " records" + bold_blue_ed
        summary_2 = [summary_2_1, 
                    iden1_op + summary_2_2 + iden_ed, 
                    iden1_op + summary_2_3 + iden_ed, 
                    iden1_op + summary_2_4 + iden_ed, 
                    iden2_op + summary_2_5 + iden_ed,
                    iden2_op + summary_2_6 + iden_ed, 
                    iden1_op + summary_2_7 + iden_ed,
                    iden1_op + summary_2_8 + iden_ed, 
                    iden2_op + summary_2_9 + iden_ed,
                    iden2_op + summary_2_10 + iden_ed, 
                    iden2_op + summary_2_11 + iden_ed, 
                    iden2_op + summary_2_12 + iden_ed]
        ##reportlab
        report_title(c,"Summary result", 1.07*inch, 10.5*inch,'#3e4444', font_size=16)
        report_context(c,summary_1, 1.0*inch, 8.0*inch, 460, 120, font_size=11, line_space=16)
        report_context(c,summary_2, 1.0*inch, 3.5*inch, 460, 300, font_size=11, line_space=16)
        report_todaypage(c,55,30,"Created on: "+today)
        report_todaypage(c,270,30,"Page " + str(page))
        c.showPage()
    
    def summary_table_1(checkpoint,df,idx,col,title_number,title_name,footnote,page,today=today.strftime("%d %b %Y")):
        ##reportlab
        if checkpoint:
            title = ""
            for i in range(len(idx)):
                if len(idx) > 1:
                    page += 1
                    if i == 0:
                        title = "<b>" + title_number + ": " + title_name + "</b>"
                        df_1 = df.loc[idx[i]:idx[i+1]]
                    elif i+1 == len(idx):
                        title = "<b>" + title_number + " (continue): " + title_name + "</b>"
                        df_1 = df.loc[idx[i]+1:]
                    else:
                        title = "<b>" + title_number + " (continue): " + title_name + "</b>"
                        df_1 = df.loc[idx[i]+1:idx[i+1]]
                else:
                    title = "<b>" + title_number + ": " + title_name + "</b>"
                    df_1 = df.loc[idx[i]:]
                df_1 = df_1.values.tolist()
                df_merge = col + df_1
                report_context(c,[title], 1.0*inch, 10.0*inch, 460, 70, font_size=11, line_space=16)
                table = df_merge
                table_draw = report_table_summary_1_v2(table)
                table_draw.wrapOn(c, 500, 300)
                h = (30-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
                table_draw.drawOn(c, 1.0*inch, (h+2.2)*inch)
                report_context(c,[footnote], 1.0*inch, 0.4*inch, 460, 100, font_size=9, line_space=12)
                report_todaypage(c,55,30,"Created on: "+today)
                report_todaypage(c,270,30,"Page " + str(page))
                c.showPage()
    
    def summary_table_2(checkpoint,df,idx,col,title_number,title_name,page,footnote_1,footnote_2=[],today=today.strftime("%d %b %Y"),x1_axis=1.0,y1_axis=1.4,width1=460,height1=100,
                                                                                                                                     x2_axis=1.0,y2_axis=0.4,width2=460,height2=100):
        ##reportlab
        if checkpoint:
            title = ""
            for i in range(len(idx)):
                if len(idx) > 1:
                    if i == 0:
                        title = "<b>" + title_number + ": " + title_name + "</b>"
                        df_1 = df.loc[idx[i]:idx[i+1]]
                    elif i+1 == len(idx):
                        title = "<b>" + title_number + " (continue): " + title_name + "</b>"
                        df_1 = df.loc[idx[i]+1:]
                    else:
                        title = "<b>" + title_number + " (continue): " + title_name + "</b>"
                        df_1 = df.loc[idx[i]+1:idx[i+1]]
                else:
                    title = "<b>" + title_number + ": " + title_name + "</b>"
                    df_1 = df.loc[idx[i]:]
                df_1 = df_1.values.tolist()
                df_merge = col + df_1
                report_context(c,[title], 1.0*inch, 10.0*inch, 460, 70, font_size=11, line_space=16)
                table = df_merge
                table_draw = report_table_summary_v2(table)
                table_draw.wrapOn(c, 500, 300)
                h = (30-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
                table_draw.drawOn(c, 1.0*inch, (h+2.2)*inch)
                if footnote_2 == []:
                    report_context(c,footnote_1, x2_axis*inch, y2_axis*inch, width2, height2, font_size=9, line_space=12)
                else:
                    report_context(c,footnote_1, x1_axis*inch, y1_axis*inch, width1, height1, font_size=9, line_space=12)
                    report_context(c,footnote_2, x2_axis*inch, y2_axis*inch, width2, height2, font_size=9, line_space=12, font_align=TA_LEFT)
                page += 1
                report_todaypage(c,55,30,"Created on: "+today)
                report_todaypage(c,270,30,"Page " + str(page))
                c.showPage()
                
    
    def table_list(df, idx, col, title_no, title_detail, page):
        title = ["<b>"+title_no+": "+title_detail+"</b>"]
        ##reportlab
        for i in range(len(idx)):
            if len(idx) > 1:
                if i == 0:
                    df_1 = df.loc[idx[i]:idx[i+1]]
                elif i+1 == len(idx):
                    title = ["<b>"+title_no+" (continue): "+title_detail+"</b>"]
                    df_1 = df.loc[idx[i]+1:]
                else:
                    title = ["<b>"+title_no+" (continue): "+title_detail+"</b>"]
                    df_1 = df.loc[idx[i]+1:idx[i+1]]
            else:
                df_1 = df.loc[idx[i]:]
            df_1 = df_1.values.tolist()
            df_merge = col + df_1
            report_context(c,title, 1.0*inch, 9.5*inch, 460, 120, font_size=11, line_space=16)
            table = df_merge
            table_draw = report_table_appendix(table)
            table_draw.wrapOn(c, 500, 300)
            h = (20-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
            table_draw.drawOn(c, 1.07*inch, (h+1.0)*inch)
            page += 1
            report_todaypage(c,55,30,"Created on: "+today.strftime("%d %b %Y"))
            report_todaypage(c,270,30,"Page " + str(page))
            c.showPage()
            
    
    def list_annexA(df, idx, col, page):
        list_noti = ["<b>Table 5: List of specimens culture positive for notifiable organisms</b>"]
        list_footnote = ["*CSF = Cerebrospinal fluid; RTS = Respiratory tract specimens; Others = Others sample types"]
        ##reportlab
        for i in range(len(idx)):
            if len(idx) > 1:
                if i == 0:
                    df_1 = df.loc[idx[i]:idx[i+1]]
                elif i+1 == len(idx):
                    list_noti = ["<b>Table 5 (continue): List of specimens culture positive for notifiable organisms</b>"]
                    df_1 = df.loc[idx[i]+1:]
                else:
                    list_noti = ["<b>Table 5 (continue): List of specimens culture positive for notifiable organisms</b>"]
                    df_1 = df.loc[idx[i]+1:idx[i+1]]
            else:
                df_1 = df.loc[idx[i]:]
            df_merge = col + (df_1.values.tolist())
            report_context(c,list_noti, 1.0*inch, 10.0*inch, 460, 70, font_size=11, line_space=16)
            table = df_merge
            table_draw = report_table_appendix_2(table)
            table_draw.wrapOn(c, 500, 300)
            h = (30-len(table))*(0.25) ####Work!!!!!!!!!!!!!!!!!!!!!!!!
            table_draw.drawOn(c, 1.0*inch, (h+2.0)*inch)
            report_context(c,list_footnote, 1.0*inch, 0.5*inch, 460, 70, font_size=9, line_space=12)
            page += 1
            report_todaypage(c,55,30,"Created on: "+today.strftime("%d %b %Y"))
            report_todaypage(c,270,30,"Page " + str(page))
            c.showPage()
            
    
    def content():
        ##content
        content = ["Summary result",]
        content_name = []
        content_page = []
        content_page_1 = [1]
        if len(summary_1) > 0:
            content_name.append("Table 1: Summary of potential contaminants")
            content_page_1.append(len(index_mark_page_summary_1))
        if len(summary_2) > 0:
            content_name.append("Table 2: Summary of notifiable antibiotic-pathogen combinations")
            content_page_1.append(len(index_mark_page_summary_2))
        if len(summary_3a) > 0:
            content_name.append("Table 3: Summary of infrequent phenotypes or potential errors in AST results based on\nthe indicators that the organisms are intrinsically resistant to an antibiotic but are reported as susceptible")
            content_page_1.append(len(index_mark_page_summary_3a))
        if len(summary_3b) > 0:
            content_name.append("Table 4: Summary of infrequent phenotypes or potential errors in AST results based on\nthe indicators that the isolates exhibit discordant AST results")
            content_page_1.append(len(index_mark_page_summary_3b))
        if len(list_alerted_bloodcon) > 0:
            content_name.append("Table 1a: List of potential contaminants")
            content_page_1.append(len(index_mark_page_bloodcon))
        if len(list_alerted_pathogen) > 0:
            content_name.append("Table 2a: List of notifiable antibiotic-pathogen combinations")
            content_page_1.append(len(index_mark_page_pathogen))
        if len(list_alerted_3a) > 0:
            content_name.append("Table 3a: List of infrequent phenotypes or potential errors in AST results based on\nthe indicators that the organisms are intrinsically resistant to an antibiotic but are reported as susceptible")
            content_page_1.append(len(index_mark_page_3a))
        if len(list_alerted_3b) > 0:
            content_name.append("Table 4a: List of infrequent phenotypes or potential errors in AST results based on\nthe indicators that the isolates exhibit discordant AST results")
            content_page_1.append(len(index_mark_page_3b))
        if len(list_alerted_A) > 0:
            content_name.append("Table 5: List of specimens culture positive for notifiable organisms")
            content_page_1.append(len(index_mark_page_A))
        content_page_2 = correct_content_page_v2(content_page_1)
        content_page_2 = ["0"+str(x) if x < 10 else str(x) for x in content_page_2]
        content = content + content_name
        content_page = content_page + content_page_2
        content_page = content_page[:4] + ["."] + ["."] + [content_page[4]] + ["."] + content_page[5:]
        ##reportlab
        report_title(c,'Content',1.07*inch, 10.5*inch,'#3e4444', font_size=16)
        report_context(c,content, 1.0*inch, 6.0*inch, 435, 300, font_size=11)
        report_context(c,content_page, 7.0*inch, 6.0*inch, 40, 300, font_size=11)
        c.showPage()
    
    def cover(hospital_name=hospital_name, country_name=country_name, spc_date_start=spc_date_start_cov, spc_date_end=spc_date_end_cov, today=today.strftime("%d %b %Y")):
        ##paragraph variable
        bold_blue_op = "<b><font color=\"#000080\">"
        bold_blue_ed = "</font></b>"
        add_blankline = "<br/>"
        ##content
        cover_1_1 = "<b>Hospital name:</b>  " + bold_blue_op + hospital_name + bold_blue_ed
        cover_1_2 = "<b>Country name:</b>  " + bold_blue_op + country_name + bold_blue_ed
        cover_1_3 = "<b>Data from:</b>"
        cover_1_4 = bold_blue_op + str(spc_date_start) + " to " + str(spc_date_end) + bold_blue_ed
        cover_1 = [cover_1_1,cover_1_2,add_blankline+cover_1_3, cover_1_4]
        cover_2_1 = "This is a detailed report for records with data indicators. This report, together with the full list in Excel format, is for users to check and validate records with notifiable bacteria, notifiable antibiotic-pathogen combinations, infrequent phenotypes or potential errors in the AST results at the local level. The identifiers listed include hospital number and specimen collection date. Users should not share or transfer this Supplementary data indictors report (in PDF and Excel formats) to any party outside of the hospital without data security management and confidential agreement."
        today = datetime.datetime.now().strftime("%d %b %Y %H:%M:%S")
        cover_2_2 = "<b>Generated on:</b>  " + bold_blue_op + today + bold_blue_ed
        cover_2 = [cover_2_1,cover_2_2]
        ##reportlab
        c.setFillColor('#FCBB42')
        c.rect(0,590,800,20, fill=True, stroke=False)
        c.setFillColor(royalblue)
        c.rect(0,420,800,150, fill=True, stroke=False)
        report_title(c,'Supplementary:',0.7*inch, 515,'white',font_size=28)
        report_title(c,'Data indicators report',0.7*inch, 455,'white',font_size=28)
        report_context(c,cover_1, 0.7*inch, 2.2*inch, 460, 240, font_size=18,line_space=26)
        report_context(c,cover_2, 0.7*inch, 0.5*inch, 460, 120, font_size=10,line_space=13)
        c.showPage()
    
    try:
        config = pd.read_excel(path + "Configuration/Configuration.xlsx")
        if check_config(config, "data_indicators_report"):
            c = canvas.Canvas(AC.CONST_PATH_REPORTWITH_PID +"Supplementary_data_indicators_report.pdf")
            cover()
            try:
                content()
            except Exception as e:
                logger.exception(e)
                pass
            page = 1
            try:
                summary_table_page1(page=page)
            except Exception as e:
                logger.exception(e)
                pass
            # page += 1
            if len(summary_1) > 0:
                try:
                    summary_table_1(checkpoint=checkpoint(summary_1_i),
                                    df=summary_1,
                                    idx=index_mark_page_summary_1,
                                    col=summary_1_col,
                                    title_number="Table 1",
                                    title_name=title_summary_1,
                                    page=page,
                                    footnote=note_summary_1)
                    page = page + len(index_mark_page_summary_1)
                except:
                    pass
            if len(summary_2) > 0:
                try:
                    summary_table_2(checkpoint=checkpoint(summary_2_i),
                                    df=summary_2,
                                    idx=index_mark_page_summary_2,
                                    col=summary_2_col,
                                    title_number="Table 2",
                                    title_name=title_summary_2,
                                    page=page,
                                    footnote_1=[note_summary_2_1],
                                    footnote_2=[note_summary_2_2], 
                                    x1_axis=1.0,y1_axis=1.4,width1=460,height1=100,
                                    x2_axis=1.0,y2_axis=0.4,width2=460,height2=70)
                    page = page + len(index_mark_page_summary_2)
                except:
                    pass
            if len(summary_3a) > 0:
                try:
                    summary_table_2(checkpoint=checkpoint(summary_3a_i),
                                    df=summary_3a,
                                    idx=index_mark_page_summary_3a,
                                    col=summary_3a_col,
                                    title_number="Table 3",
                                    title_name=title_summary_3,
                                    page=page,
                                    footnote_1=[note_summary_3_1])
                    page = page + len(index_mark_page_summary_3a)
                except:
                    pass
    
            if len(summary_3b) > 0:
                try:
                    summary_table_2(checkpoint=checkpoint(summary_3b_i),
                                    df=summary_3b,
                                    idx=index_mark_page_summary_3b,
                                    col=summary_3b_col,
                                    title_number="Table 4",
                                    title_name=title_summary_4,
                                    page=page,
                                    footnote_1=[note_summary_3_1, note_summary_3_2, note_summary_3_3, note_summary_3_4, note_summary_3_5, note_summary_3_6], 
                                    x2_axis=1.0,y2_axis=0.7,width2=460,height2=270)
                    page = page + len(index_mark_page_summary_3b)
                except:
                    pass
    
            if len(list_alerted_bloodcon) > 0:
                try:
                    table_list(df=list_alerted_bloodcon, idx=index_mark_page_bloodcon, col=list_alerted_bloodcon_col, title_no="Table 1a", title_detail="List of potential contaminants", page=page)
                    page = page + len(index_mark_page_bloodcon) + 1
                except:
                    pass
            if len(list_alerted_pathogen) > 0:
                try:
                    table_list(df=list_alerted_pathogen, idx=index_mark_page_pathogen, col=list_alerted_pathogen_col, title_no="Table 2a", title_detail="List of notifiable antibiotic-pathogen combinations", page=page)
                    page = page + len(index_mark_page_pathogen) + 1
                except:
                    pass
            if len(list_alerted_3a) > 0:
                try:
                    table_list(df=list_alerted_3a, idx=index_mark_page_3a, col=list_alerted_3a_col, title_no="Table 3a", title_detail="List of infrequent phenotypes or potential errors in AST results based on the indicators that the organisms are intrinsically resistant to an antibiotic but are reported as susceptible", page=page)
                    page = page + len(index_mark_page_3a) + 1
                except:
                    pass
            if len(list_alerted_3b) > 0:
                try:
                    table_list(df=list_alerted_3b, idx=index_mark_page_3b, col=list_alerted_3b_col, title_no="Table 4a", title_detail="List of infrequent phenotypes or potential errors in AST results based on the indicators that the organisms with discordant AST results", page=page)
                    page = page + len(index_mark_page_3b) + 1
                except:
                    pass
            if checkpoint(summary_list_A_i) and len(list_alerted_A) > 0:
                try:
                    list_annexA(df=list_alerted_A, idx=index_mark_page_A, col=list_alerted_A_col, page=page)
                except:
                    pass
            c.save()
        else:
            pass
    except Exception as e:
        logger.exception(e)
        pass