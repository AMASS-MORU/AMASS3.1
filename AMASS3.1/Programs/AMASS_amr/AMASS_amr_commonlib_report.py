#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.

# Created on 20th April 2022 (V2.0)
import math as m
import gc
import pandas as pd #for creating and manipulating dataframe
import matplotlib.pyplot as plt #for creating graph (pyplot)
#import matplotlib #for importing graph elements
import numpy as np #for creating arrays and calculations
import seaborn as sns #for creating graph
from pathlib import Path #for retrieving input's path
from datetime import date #for generating today date
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT #for setting paragraph style
from reportlab.platypus import Table,Frame #for plotting graph and tables
#from reportlab.lib.colors import * #for importing color palette
#from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib import colors #for importing color palette
#from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
"""
# Core pages
pagenumber_intro = ["1","2"]
pagenumber_ava_1 = ["3","4"] #section1; 2 pages
pagenumber_ava_2 = ["5","6","7","8","9","10","11"] #section2; 7 pages

# Available hospital_admission_data.xlsx
pagenumber_ava_3 = ["12","13","14","15","16","17","18","19","20","21","22","23"] #section3; 12 pages
pagenumber_ava_4 = ["24","25","26"] #section4; 3 pages
pagenumber_ava_5 = ["27","28","29","30","31"] #section1; 5 pages
pagenumber_ava_6 = ["32","33","34","35","36","37"] #section6; 6 pages
pagenumber_ava_annexA = ["38","39","40"] #annexA; 3 pages
pagenumber_ava_annexB = ["41","42"] #annexB; 2 pages
pagenumber_ava_other = ["43","44","45","46","47"] #complement; 6 pages
"""
CONST_COLORPALETTE_1 = ['rebeccapurple','darkorange','firebrick','dodgerblue','saddlebrown','yellowgreen','palevioletred','darkkhaki']

# Added in verions 3.0------------------------------------------------------------
#CONST
##### Cal figure height for section6
def cal_sec6_fig_height(irow):
    if irow <= 2:
        return 1
    elif irow <= 3:
        return 1.4 
    elif irow <= 5:
        return 2 
    else:
        return 3.5
##### Cal figure height for section4,5
def cal_sec4_fig_height(irow):
    if irow <= 6:
        return 3.0
    elif irow <= 9:
        return 6.0 #default height
    else:
        return 7.0 
##### Cal figure height for section2 and section3
def cal_sec2and3_fig_height(iatbrow):
    if iatbrow <= m.ceil(AC.CONST_MAX_ATBCOUNTFITHALFPAGE/2):
        return 2.62
    elif iatbrow <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE:
        return 3.5 #default height
    elif iatbrow <= AC.CONST_MAX_ATBCOUNTFITHALFPAGE*1.2:
        return 4.5 #default height
    else:
        return 5.47 
def count_atbperorg(raw_df,org_full,origin="", org_col="Organism", drug_col="Antibiotic",  ns_col="Non-susceptible(N)", total_col="Total(N)"):
    if origin == "": #section2
        sel_df = raw_df.loc[raw_df[org_col]==org_full].set_index(drug_col).fillna(0)
    else: #section3
        sel_df = raw_df.loc[(raw_df[org_col]==org_full) & (raw_df['Infection_origin']==origin)].set_index(drug_col).fillna(0)
    return len(sel_df)
def get_atbnote(list_curnote):
    try:
        sallnote = ""
        for scurnote in list_curnote:
            scurnote = scurnote.strip()
            s = AC.dict_atbnote.get(scurnote, "")
            if s != "":
                sallnote = sallnote + ("" if sallnote == "" else "; ") + s
        return sallnote
    except Exception as e:
        print(e)
        return ""
        pass
def get_atbnoteperorg(raw_df,org_full,origin="",origin_col="Infection_origin", org_col="Organism", drug_col="Antibiotic",  ns_col="Non-susceptible(N)", total_col="Total(N)"): 
    if origin == "": #section2
        list_atbnote = raw_df.loc[raw_df[org_col]==org_full][drug_col].values.tolist()
    else: #section3
        list_atbnote = raw_df.loc[(raw_df[org_col]==org_full) & (raw_df[origin_col]==origin)][drug_col].values.tolist()
    return list_atbnote
def get_atbnoteper_priority_pathogen(raw_df,origin="",origin_col="Infection_origin",pathogen_col="Priority_pathogen"):
    satbnote = ""
    list_atbnote = []
    df =raw_df.copy(deep=True)
    if origin != "": #section4
        df = df.loc[(df[origin_col]==origin)]
    for ind in df.index:
        for spt in AC.dict_atbnote_sec4_5.keys():
            sdfpt = df[pathogen_col][ind].lower()
            sptl = spt.lower()
            if sptl in sdfpt:
                if not (sptl in list_atbnote):
                    list_atbnote = list_atbnote + [sptl]
                    #satbnote = satbnote + " " + AC.dict_atbnote_sec4_5[spt]
                    spt = spt.strip()
                    s = AC.dict_atbnote_sec4_5.get(spt, "")
                    if s != "":
                        satbnote = satbnote + ("" if satbnote == "" else "; ") + s
    return satbnote
def get_atbnoteper_priority_pathogen_sec6(raw_df,origin="",origin_col="Infection_origin",pathogen_col="Antibiotic"):
    list_atbnote = []
    df =raw_df.copy(deep=True)
    if origin != "": #section4
        df = df.loc[(df[origin_col]==origin)]
    for ind in df.index:
        for spt in AC.dict_atbnote_sec6.keys():
            sdfpt = df[pathogen_col][ind].lower()
            sptl = spt.lower()
            if sptl in sdfpt:
                if not (sptl in list_atbnote):
                    list_atbnote = list_atbnote + [spt]
    return list_atbnote
def get_atbnote_sec6(list_atbnote):
    snote = ""
    for scurnote in list_atbnote:
        if scurnote not in snote:
            scurnote = scurnote.strip()
            s = AC.dict_atbnote_sec6.get(scurnote, "")
            if s != "":
                snote = snote + ("" if snote == "" else "; ") + s
            #snote = snote + AC.dict_atbnote_sec6[scurnote]
    return snote   
def create_table_nons_V3(raw_df, org_full, origin="", iMODE=AC.CONST_VALUE_MODE_AST):
    #Selecting rows by organism and parsing table for PDF
    #raw_df : raw dataframe is opened from AC.CONST_FILENAME_sec2_amr_i
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_col : column name of organisms
    #drug_col : column name of antibiotics
    #ns_col : column name of number of non-susceptible
    #total_col : column name of total number of patient
    org_col ="Organism"
    drug_col = 'Antibiotic'
    ns_col = "Non-susceptible(N)"
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        ns_col = "Resistant(N)"
    total_col="Total(N)"
    upper_col = 'upper95CI(%)*'
    lower_col = 'lower95CI(%)*'
    plabel = 'Proportion of\n NS isolates (n)'
    if iMODE != AC.CONST_VALUE_MODE_AST:
        ns_col = "Resistant(N)"
        total_col="Total(N)"
        upper_col = 'Resistant-upper95CI(%)*'
        lower_col = 'Resistant-lower95CI(%)*'
        #plabel = 'Proportion of\n resistant isolates (n)'
        plabel = 'Proportion of R'
    raw_df[drug_col] = raw_df[drug_col].str.strip()
    #for ind in raw_df.index:
    #    print("'" + raw_df['Antibiotic'][ind] + "'")
    if origin == "": #section2
        sel_df = raw_df.loc[raw_df[org_col]==org_full].set_index(drug_col).fillna(0)
    else: #section3
        sel_df = raw_df.loc[(raw_df[org_col]==org_full) & (raw_df['Infection_origin']==origin)].set_index(drug_col).fillna(0)
    sel_df["%"] = round(sel_df[ns_col]/sel_df[total_col]*100,1).fillna(0)
    #sel_df_1 = correct_digit(df=sel_df,df_col=["lower95CI(%)*","upper95CI(%)*","%"])
    sel_df_1 = correct_digit(df=sel_df,df_col=[lower_col,upper_col,"%"])
    
    sel_df_1["% NS (n)"] = sel_df_1["%"].astype(str) + "% (" + sel_df_1[ns_col].astype(str) + "/" + sel_df_1[total_col].astype(str) + ")"
    sel_df_1["95% CI"]  = sel_df_1[lower_col].astype(str) + "% - " + sel_df_1[upper_col].astype(str) + "%"
    sel_df_1 = sel_df_1.loc[:,["% NS (n)", "95% CI"]].reset_index().rename(columns={drug_col:'Antibiotic agent',"% NS (n)":plabel})
    sel_df_1 = sel_df_1.replace("0% (0/0)","NA").replace("0% - 0%","-")
    
    col = pd.DataFrame(list(sel_df_1.columns)).T
    col.columns = list(sel_df_1.columns)
    #return col.append(sel_df_1)
    return pd.concat([col,sel_df_1],ignore_index = True)
def create_graph_nons_v3(raw_df, org_full, org_short, palette, drug_col, iMODE,ifig_H, origin=""):
    #Creating graph for PDF
    #raw_df : raw dataframe is opened from AC.CONST_FILENAME_sec2_amr_i
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_short : short name of organisms for using to retrieve rows by names ex.s_aureus
    #org_col : column name of organisms
    #drug_col : column name of antibiotics
    #perc_col : column name of %Non-susceptible
    #upper_col : column name of upperCI
    #lower_col : column name of lowerCT
    #print("FOR ORG " + org_full)
   
    perc_col = 'Non-susceptible(%)'
    if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
        perc_col = 'Resistant(%)'
    upper_col = 'upper95CI(%)*'
    lower_col = 'lower95CI(%)*'
    xlabel = '*Proportion of NS isolates (%)'
    if iMODE != AC.CONST_VALUE_MODE_AST:
        perc_col = 'Resistant(%)'
        upper_col = 'Resistant-upper95CI(%)*'
        lower_col = 'Resistant-lower95CI(%)*'
        xlabel = '*Proportion of R (%)'
    raw_df[drug_col] = raw_df[drug_col].str.rjust(AC.CONST_MAXNUMCHAR_ATBNAME," ")
    if origin == "": #section2
        sel_df = raw_df.loc[raw_df['Organism']==org_full].set_index(drug_col).fillna(0)
    else: #section3
        sel_df = raw_df.loc[(raw_df['Organism']==org_full) & (raw_df['Infection_origin']==origin)].set_index(drug_col).fillna(0)
    #print(sel_df)
    plt.figure(figsize=(7,ifig_H))
    sns.barplot(data=sel_df.loc[:,perc_col].to_frame().T, palette=palette,orient='h',capsize=.2)
    for drug in sel_df.index:
        ci_lo=sel_df.loc[drug,lower_col]
        ci_hi=sel_df.loc[drug,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.xlabel(xlabel,fontsize=14)
    plt.ylabel('', fontsize=11)
    sns.despine(left=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=19)
    plt.yticks(fontname='sans-serif',style='normal',fontsize=19)
    plt.tick_params(top=False, bottom=True, left=False, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    if origin == "":
        plt.savefig(AC.CONST_PATH_RESULT + 'Report2_AMR_' + org_short + '.png', format='png',dpi=180,transparent=True)
    else:
        plt.savefig(AC.CONST_PATH_RESULT + 'Report3_AMR_' + org_short + '_' + origin + '.png', format='png',dpi=180,transparent=True)
    plt.close()
    plt.clf
def prepare_section2_table_for_reportlab_V3(df_org, df_pat, lst_org,lst_org_format, checkpoint_sec2,dict_orgcatwithatb):
    ##Preparing section2_table for ploting in pdf
    #df: raw section2_table
    #checkpoint_hosp: checkpoint of section2_result
    #return value: lst of dataframe "section2_table"
    d_org_core = {}
    for sorgkey in dict_orgcatwithatb:
        ocurorg = dict_orgcatwithatb[sorgkey]
        d_org_core[sorgkey] = ocurorg[2]
    df_org = df_org.set_index("Organism").rename(d_org_core)
    df_pat = df_pat.set_index("Organism").rename(d_org_core)
    if checkpoint_sec2:
        #Creating table for page 6
        df_merge = pd.merge(df_org.astype(str), df_pat.astype(str), on="Organism", how="outer").fillna("0").loc[lst_org]
        df_merge_sum = pd.merge(df_org, df_pat, on="Organism", how="outer").fillna(0)
        if "organism_no_growth" in df_merge_sum.index:
            df_merge_sum = df_merge_sum.drop(index=["organism_no_growth"])
        else:
            pass
        #Reformetting Organism name
        d_org_fmt = {}
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=10,alignment=TA_LEFT)
        for i in range(len(lst_org)):
            d_org_fmt[lst_org[i]] = Paragraph("<b>" + lst_org_format[i] + "</b>",style_summary)
        df_merge = df_merge.rename(index=d_org_fmt)
        #Adding Total
        df_merge.loc["Total:","Number_of_blood_specimens_culture_positive_for_the_organism"] = round(df_merge_sum["Number_of_blood_specimens_culture_positive_for_the_organism"].sum())
        df_merge.loc["Total:","Number_of_blood_specimens_culture_positive_deduplicated"]     = round(df_merge_sum["Number_of_blood_specimens_culture_positive_deduplicated"].sum())
        df_merge = df_merge.reset_index().rename(columns={"Number_of_blood_specimens_culture_positive_for_the_organism":"Number of records\nof blood specimens\nculture positive\nfor the organism", 
                                                            "Number_of_blood_specimens_culture_positive_deduplicated":"**Number of patients with\nblood culture positive\nfor the organism\n(de−duplicated)"})
        #Preparing to list
        lst_col = [list(df_merge.columns)]
        lst_df = df_merge.values.tolist()
        lst_df = lst_col + lst_df
    else:
        #Reformetting Organism name
        d_org_fmt = {}
        style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=11,alignment=TA_LEFT)
        for i in range(len(lst_org)):
            d_org_fmt[lst_org[i]] = Paragraph("<b>" + lst_org_format[i] + "</b>",style_summary)
        df_merge = pd.DataFrame(index=lst_org + ["Total:"], 
                                columns=["Number of records\nof blood specimens\nculture positive\nfor the organism", 
                                        "**Number of patients with\nblood culture positive\nfor the organism\n(de−duplicated)"])
        lst_df = df_merge.rename(index=d_org_fmt).fillna("NA").values.tolist()
    return lst_df
def prepare_section3_table_for_reportlab_V3(df_pat, lst_org, lst_org_format):
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=11,alignment=TA_LEFT)
    for i in range(len(lst_org)):
        df_pat.loc[df_pat['Organism']==lst_org[i], 'Organism'] =  Paragraph("<b>" + lst_org_format[i] + "</b>",style_summary)
    df_pat.loc[df_pat['Organism']=='Total', 'Organism'] = Paragraph("<b>" + 'Total:' + "</b>",style_summary)
    return df_pat
def prepare_section6_mortality_table_for_reportlab_V3(df_mor, lst_org, lst_org_format):
    ##Preparing section6_mortality_table for reportlab; page 33
    df_mor_com = create_table_perc_mortal_eachorigin_V3(df_mor,'Organism','Infection_origin','Number_of_deaths','Total_number_of_patients','Community-origin','Mortality in patients with\nCommunity−origin BSI')
    df_mor_hos = create_table_perc_mortal_eachorigin_V3(df_mor,'Organism','Infection_origin','Number_of_deaths','Total_number_of_patients','Hospital-origin', 'Mortality in patients with\nHospital−origin BSI')
    df_mor_all = df_mor_com.merge(df_mor_hos, how="inner", left_on='Organism', right_on='Organism',suffixes=("", "_hos"))
    style_summary = ParagraphStyle('normal',fontName='Helvetica',fontSize=11,alignment=TA_LEFT)
    for i in range(len(lst_org)):
        df_mor_all.loc[df_mor_all['Organism']==lst_org[i], 'Organism'] =  Paragraph("<b>" + lst_org_format[i] + "</b>",style_summary)
    df_mor_all.loc[df_mor_all['Organism']=='Total:', 'Organism'] = Paragraph("<b>" + 'Total:' + "</b>",style_summary)
    df_mor_dis = pd.DataFrame(list(df_mor_all.columns),index=df_mor_all.columns).T
    #sec3_pat_val = sec3_col.append(sec3_pat).drop(columns=["Organism"]).values.tolist()
    return df_mor_dis.append(df_mor_all).values.tolist()
def create_table_perc_mortal_eachorigin_V3(df,org_col,ori_col,mortal_col,total_col,origin,col_displayval): 
    df = df.loc[df[ori_col]==origin,:].astype({mortal_col:'int32',total_col:'int32'})
    df_amr = df[[org_col,mortal_col,total_col]]
    df_amr.loc["Total",org_col] = "Total:"
    msum = df_amr[mortal_col].sum()
    tsum = df_amr[total_col].sum()
    df_amr.loc["Total",mortal_col] = msum
    df_amr.loc["Total",total_col] = tsum
    df_amr['perc_mortal'] = round((df_amr[mortal_col]/df_amr[total_col]*100),0).fillna(0)
    #df_amr_1 = correct_digit(df=df_amr, df_col=["perc_mortal",mortal_col,total_col])
    df_amr[col_displayval] = (df_amr['perc_mortal'].astype(int).astype(str) + '% (' + df_amr[mortal_col].astype(int).astype(str) + '/' + df_amr[total_col].astype(int).astype(str) + ')').replace('0% (0/0)','NA')
    #return df_amr[[org_col,col_displayval]]
    return df_amr[[org_col,col_displayval]]
def create_graph_surveillance_V3(df_raw, lst_org, prefix, text_work_drug="N", freq_col="frequency_per_tested", upper_col="frequency_per_tested_uci", lower_col="frequency_per_tested_lci",ifig_H=12):
    #Creating graph for PDF
    #raw_df : raw dataframe is opened from AC.CONST_FILENAME_sec2_amr_i
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #upper_col : column name of upperCI
    #lower_col : column name of lowerCT
    if text_work_drug == "Y":
        df_raw = df_raw.set_index('Priority_pathogen')
        try:
            palette = create_graphpalette_sec4_5(df_raw,'Organism',CONST_COLORPALETTE_1)
        except:
            palette = ['rebeccapurple','darkorange','firebrick','dodgerblue','saddlebrown','saddlebrown','yellowgreen','yellowgreen','palevioletred','darkkhaki']
            #print("Warning : error get color palette for sec4,5 - pathogen")
    else:
        try:
            palette = create_graphpalette_sec4_5(df_raw,'Organism',CONST_COLORPALETTE_1)
        except:
            palette = ['rebeccapurple','darkorange','firebrick','dodgerblue','saddlebrown','yellowgreen','palevioletred','darkkhaki']
            #print("Warning : error get color palette for sec4,5 - organism")
        df_raw = df_raw.set_index('Organism')
    plt.figure(figsize=(7,ifig_H))
    sns.barplot(data=df_raw.loc[:,freq_col].to_frame().T,palette=palette,orient='h',capsize=.2)
    for idx in df_raw.index:
        ci_lo=df_raw.loc[idx,lower_col]
        ci_hi=df_raw.loc[idx,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[idx,idx],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[idx,idx],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.locator_params(axis="x", nbins=7) #set number of bins for x-axis
    plt.xlim(0,round(df_raw[upper_col].max())+50)
    plt.xlabel('*Frequency of infection\n(per 100,000 tested patients)',fontsize=14)
    plt.ylabel('', fontsize=11)
    # sns.despine(left=True)
    sns.despine(top=True,right=True) #set REMOVED border line
    plt.xticks(fontname='sans-serif',style='normal',fontsize=18)
    plt.yticks(np.arange(len(lst_org)),lst_org,fontname='sans-serif',style='normal',fontsize=18)
    # plt.yticks("",fontname='sans-serif',style='normal',fontsize=20)
    plt.tight_layout()
    plt.savefig(AC.CONST_PATH_RESULT + prefix + '.png', format='png',dpi=180,transparent=True)
    plt.close()
    plt.clf
def create_graph_mortal_V3(df, organism, origin, prefix, org_col="Organism", ori_col="Infection_origin", drug_col="Antibiotic", perc_col="Mortality (n)", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    ##Creating graph of mortality
    df_1 = df.loc[df[org_col]==organism,:]
    df_1 = df_1.loc[df[ori_col]==origin,:].replace(regex=["3GC-NS"],value="3GC-NS**").replace(regex=["3GC-S"], value="3GC-S***").replace(regex=["3GC-R"], value="3GC-R**").set_index(drug_col)
    ifig_H =2*cal_sec6_fig_height(len(df_1))
    plt.figure(figsize=(7,ifig_H))
    sns.barplot(data=df_1.loc[:,perc_col].to_frame().astype(int).T, palette=['darkorange','darkorange'],orient='h',capsize=.2)
    for drug in df_1.index:
        ci_lo=df_1.loc[drug,lower_col]
        ci_hi=df_1.loc[drug,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[drug,drug],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.xlabel('*Mortality (%)',fontsize=14)
    plt.ylabel('', fontsize=11)
    sns.despine(left=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=20)
    plt.yticks(fontname='sans-serif',style='normal',fontsize=20)
    plt.tick_params(top=False, bottom=True, left=False, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    plt.savefig(AC.CONST_PATH_RESULT + prefix + '.png', format='png',dpi=180,transparent=True)
    plt.close()
    plt.clf
#---------------------------------------------------------------------------------

def check_config(df_config, str_process_name):
    #Checking process is either able for running or not
    #df_config: Dataframe of config file
    #str_process_name: process name in string fmt
    #return value: Boolean; True when parameter is set "yes", False when parameter is set "no"
    config_lst = df_config.iloc[:,0].tolist()
    result = True
    try:
        if df_config.loc[config_lst.index(str_process_name),"Setting parameters"] == "yes":
            result = True
        else:
            result = False
    except:
        pass
    return result

def checkpoint(str_filename):
    #Checking file available
    #return : boolean ; True when file is available; False when file is not available
    return Path(str_filename).is_file()

def prepare_section1_table_for_reportlab(df, checkpoint_hosp):
    ##Preparing section1_table for ploting in pdf
    #df: raw section1_table
    #checkpoint_hosp: checkpoint of hospital_admission_data
    #return value: lst of dataframe "section1_table"
    df_sum = df.set_index(df.columns[0])
    df = df.set_index(df.columns[0]).fillna(0).astype(int).astype(str)
    df.at["Total",df.columns[0]] = round(df_sum.iloc[:,0].sum(skipna=True))
    df = df.reset_index()
    if checkpoint_hosp:
        df.loc[df["Month"]=="Total","Number_of_hospital_records_in_hospital_admission_data_file"] = round(df_sum.iloc[:,1].sum(skipna=True))
    else:
        df["Number_of_hospital_records_in_hospital_admission_data_file"] = ""
        df.loc[df["Month"]=="Total","Number_of_hospital_records_in_hospital_admission_data_file"] = "NA"
    df = df.rename(columns={"Number_of_specimen_in_microbiology_data_file":"Number of specimen\ndata records in\nmicrobiology_data file", "Number_of_hospital_records_in_hospital_admission_data_file":"Number of admission\ndata records in\nhospital_admission_data file"})
    
    lst_col = [list(df.columns)]
    lst_df = df.values.tolist()
    lst_df = lst_col + lst_df
    return lst_df

def correct_digit(df=pd.DataFrame(),df_col=[]):
    df_new = df.copy().astype(str)
    for idx in df.index:
        for col in df_col:
            if float(df.loc[idx,col]) < 0.05:
                df_new.at[idx,col] = "0"
            elif float(df.loc[idx,col]) >= 0.95:
                df_new.at[idx,col] = str(int(round(df.loc[idx,col])))
            else:
                pass
    return df_new

def create_num_patient(raw_df, org_full, org_col, ori_col=""):
    #Retrieving number of positive patient for each organism
    #raw_df : raw dataframe is opened from AC.CONST_FILENAME_sec2_pat_i
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_col : column name of organisms
    if ori_col == "": #section2
        temp_table = raw_df.loc[raw_df[org_col]==org_full].values.tolist() #[['Staphylococcus aureus', 100]]
    else: #section3
        temp_table = raw_df.loc[:,['Organism',ori_col]]
        temp_table = temp_table.loc[raw_df[org_col]==org_full].values.tolist() #[['Staphylococcus aureus', 100]]
    return temp_table[0][1] #100
def create_graphpalette_sec4_5(df,key_col,list_palettetemplate):
    list_palette = []
    cur_key_colval = ""
    cur_pidx = -1
    for idx in df.index:
        p = 'darkorange'
        try:
            if df.loc[idx,key_col] == cur_key_colval:
                p = list_palettetemplate[cur_pidx]
            else:
                if (cur_pidx < 0) or (cur_pidx >= (len(list_palettetemplate)-1)) :
                    cur_pidx =0
                else:
                    cur_pidx = cur_pidx + 1
                p = list_palettetemplate[cur_pidx]
        except Exception as e:
            print("Warning : error get palette : " + str(e))
            pass
        list_palette.append(p)
    return list_palette    
def create_graphpalette(numer_df, numer_col, org_col, org_full, denom_num, cutoff=70.0, origin=""):
    #Function for creating color palette of each organism based on 70.0 cutoff ratio (default)
    #Lower than cutoff : color gainsboro (very light grey)
    #Equal or higher than cutoff : color darkorange
    #numer_df : raw dataframe is opened from AC.CONST_FILENAME_sec2_amr_i
    #numer_col : column name that will be used as numerator
    #org_col : column name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #org_full : full name of organisms for using to retrieve rows by names ex.Staphylococcus aureus
    #denom_num : number that will be used as denominator
    #cutoff : default is 70.0% (percentage; 100%)
    if origin == "": #section2
        sel_df = numer_df.loc[numer_df[org_col]==org_full]
    else:
        sel_df = numer_df.loc[(numer_df[org_col]==org_full) & (numer_df['Infection_origin']==origin)]
    palette = []
    for idx in sel_df.index:
        perc = sel_df.loc[idx,numer_col]/denom_num*100
        if perc < cutoff:
            palette.append('gainsboro')
        else:
            palette.append('darkorange')
    return palette

def create_table_surveillance_1(df_raw, lst_org, text_work_drug="N", freq_col="frequency_per_tested", upper_col="frequency_per_tested_uci", lower_col="frequency_per_tested_lci"):
    df_merge = pd.concat([pd.DataFrame(lst_org,columns=["Organism_fmt"]),df_raw],axis=1)
    
    if text_work_drug == "N":
        df_merge = df_merge.drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Organism"})
    else:
        df_merge = df_merge.drop(columns=["Organism","Priority_pathogen"]).rename(columns={"Organism_fmt":"Organism"})
    for c in [freq_col,upper_col,lower_col]:
        df_merge[c+'_1'] = df_merge[c].astype(int)
        for idx in df_merge.index:
            if df_merge.loc[idx,c+'_1'] >= 0.05: #rounding up values of freq, lci, and uci columns
                df_merge.loc[idx,c+'_1'] = df_merge.loc[idx,c+'_1'] + 1
            else:
                df_merge.loc[idx,c+'_1'] = 0
    df_merge["*Frequency (95% CI)"] = df_merge[freq_col+'_1'].astype(str) + \
                                    "\n (" + df_merge[lower_col+'_1'].astype(str) + "-" + \
                                    df_merge[upper_col+'_1'].astype(str) + ")" #creating '*Frequency (95% CI)' columns
    df_merge["*Frequency (95% CI)"] = df_merge["*Frequency (95% CI)"].replace("0\n (0-0)","NA")
    df_merge_1 = df_merge.loc[:,["Organism","*Frequency (95% CI)"]]  #content
    if text_work_drug == "N":
        df_merge_1 = df_merge_1.rename(columns={"Organism":"Pathogens","*Frequency (95% CI)":"*Frequency of infection\n(per 100,000 tested patients;\n95% CI)"})
    else:
        if AC.CONST_MODE_R_OR_AST  != AC.CONST_VALUE_MODE_AST:
            df_merge_1 = df_merge_1.rename(columns={"Organism":"Resistant\n(NS) pathogens","*Frequency (95% CI)":"*Frequency of infection\n(per 100,000 tested patients;\n95% CI)"})
        else:
            df_merge_1 = df_merge_1.rename(columns={"Organism":"Non-susceptible\n(NS) pathogens","*Frequency (95% CI)":"*Frequency of infection\n(per 100,000 tested patients;\n95% CI)"})
    lst_col = [list(df_merge_1.columns)]
    lst_df = df_merge_1.values.tolist()
    lst_df = lst_col + lst_df
    return lst_df

def prepare_section6_mortality_table(df_amr, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    ##Preparing mortality table; page34-37
    df_amr["Mortality (n)"] = round(df_amr[death_col]/df_amr[total_col]*100,1).fillna(0)
    df_amr_1 = correct_digit(df=df_amr, df_col=["Mortality (n)", lower_col, upper_col])
    # df_amr_1 = df_amr.copy().astype(str)
    # for idx in df_amr.index:
    #     for col in ["Mortality (n)", lower_col, upper_col]:
    #         if float(df_amr.loc[idx,col]) < 0.05:
    #             df_amr_1.at[idx,col] = "0"
    #         elif float(df_amr.loc[idx,col]) >= 0.95:
    #             df_amr_1.at[idx,col] = str(int(round(df_amr.loc[idx,col])))
    #         else:
    #             pass
    df_amr_1["Mortality (n)"] = (df_amr_1["Mortality (n)"]+"% ("+df_amr_1[death_col].astype(str)+"/"+df_amr_1[total_col].astype(str)+")").replace(regex=["0% (0/0)"],value="NA")
    df_amr_1["95% CI"] = (df_amr_1[lower_col].astype(str)+'% - '+df_amr_1[upper_col].astype(str)+"%").replace(regex=["0.0% - 0.0%"],value="-")
    df_amr_1["Antibiotic"] = df_amr_1["Antibiotic"].replace(regex=["3GC-NS"], value="3GC-NS**").replace(regex=["3GC-S"], value="3GC-S***").replace(regex=["3GC-R"], value="3GC-R**")
    return df_amr_1.loc[:,["Organism","Antibiotic","Infection_origin","Mortality (n)", "95% CI"]].rename(columns={'Antibiotic':'Type of pathogen'})

def create_table_mortal(df, organism, origin, ori_col="Infection_origin", org_col="Organism"):
    df_1 = df.loc[df[ori_col]==origin,:]
    df_2 = df_1.loc[df_1[org_col]==organism].drop(columns=[org_col,ori_col])
    df_col = pd.DataFrame(df_2.columns,index=df_2.columns).T #column name
    #df_3 = df_col.append(df_2).replace("0% (0/0)","NA").replace('0% - 0%',"-") #column name + content
    df_3 = pd.concat([df_col,df_2]).replace("0% (0/0)","NA").replace('0% - 0%',"-")
    return df_3

def prepare_section6_numpat_dict(df_mor, origin, origin_col="Infection_origin", total_col="Total_number_of_patients"):
    ##Preparing dataframe of section6_numpat for page 3-6
    df_mor = df_mor.loc[df_mor[origin_col]==origin,["Organism",total_col]] #.drop(columns=["Antibiotic","Infection_origin","Number_of_deaths","Mortality","Mortality_lower_95ci","Mortality_upper_95ci"]).set_index('Organism').astype(int)
    return (df_mor.groupby(['Organism'],sort=False).sum())

def prepare_annexA_numpat_table_for_reportlab(df_pat, lst_org):
    annexA_org_page2 = pd.DataFrame(lst_org,columns=["Organism_fmt"])
    df_pat = df_pat.rename(columns={"Total_number_of_patients":"Total number\nof patients*", 
                                    "Number_of_patients_with_blood_positive_deduplicated":"Blood",
                                    "Number_of_patients_with_csf_positive_deduplicated":"CSF",
                                    "Number_of_patients_with_genitLal_swab_positive_deduplicated":"Genital\nswab", 
                                    "Number_of_patients_with_rts_positive_deduplicated":"RTS", 
                                    "Number_of_patients_with_stool_positive_deduplicated":"Stool", 
                                    "Number_of_patients_with_urine_positive_deduplicated":"Urine", 
                                    "Number_of_patients_with_others_positive_deduplicated":"Others"})
    df_pat_2 = pd.concat([annexA_org_page2,df_pat],axis=1).drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Pathogens"})
    df_pat_2.iloc[-1,0] = "Total"
    return [list(df_pat_2.columns)] + df_pat_2.values.tolist()

def prepare_annexA_mortality_table_for_reportlab(df_mor, lst_org, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    df_mor["Mortality(%)"] = round(df_mor[death_col]/df_mor[total_col]*100,1).fillna(0)
    # df_mor_1 = df_mor.copy().astype(str)
    df_mor_1 = correct_digit(df=df_mor,df_col=["Mortality(%)",upper_col,lower_col])
    df_mor_1[["Mortality (n)","95% CI"]] = ""
    df_mor_1["Mortality (n)"] = (df_mor_1["Mortality(%)"]+"% ("+df_mor_1[death_col].astype(str)+"/"+df_mor_1[total_col].astype(str)+")").replace(regex=["0% (0/0)"],value="NA")
    df_mor_1['95% CI'] = (df_mor_1[lower_col].astype(str)+'% - '+df_mor_1[upper_col].astype(str)+"%").replace(regex=["0.0% - 0.0%"],value="-")
    df_mor_1 = df_mor_1.loc[:,["Organism","Mortality (n)","95% CI"]].replace("0% (0/0)","NA").replace("0% - 0%","-")
    df_mor_2 = pd.concat([lst_org,df_mor_1],axis=1).drop(columns=["Organism"]).rename(columns={"Organism_fmt":"Pathogens"}) #Adding orgaism_fmt to table
    return [list(df_mor_2.columns)] + df_mor_2.values.tolist()

#import numpy as np
def create_annexA_mortality_graph(df_mor, lst_org, death_col="Number_of_deaths", total_col="Total_number_of_patients", lower_col="Mortality_lower_95ci", upper_col="Mortality_upper_95ci"):
    df_mor["Mortality(%)"] = round(df_mor[death_col]/df_mor[total_col]*100,1).fillna(0)
    df_mor = df_mor.set_index('Organism')
    palette = ['darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange','darkorange']
    plt.figure(figsize=(5.0,10))
    sns.barplot(data=df_mor.loc[:,'Mortality(%)'].astype(float).to_frame().T, palette=palette,orient='h',capsize=.2)
    for org in df_mor.index:
        ci_lo=df_mor.loc[org,lower_col]
        ci_hi=df_mor.loc[org,upper_col]
        if ci_lo == 0 and ci_hi == 0:
            plt.plot([ci_lo,ci_hi],[org,org],'|-',color='black',markersize=12,linewidth=0,markeredgewidth=0)
        else:
            plt.plot([ci_lo,ci_hi],[org,org],'|-',color='black',markersize=12,linewidth=3,markeredgewidth=3)
    plt.xlim(0, 100)
    plt.ylabel('', fontsize=10)
    sns.despine(top=True, right=True)
    plt.xticks(fontname='sans-serif',style='normal',fontsize=14)
    plt.yticks(np.arange(len(lst_org)),lst_org,fontname='sans-serif',style='normal',fontsize=14)
    plt.tick_params(top=False, bottom=True, left=True, right=False,labelleft=True, labelbottom=True)
    plt.tight_layout()
    plt.savefig(AC.CONST_PATH_RESULT + 'AnnexA_mortality.png', format='png',dpi=300,transparent=True)
    plt.close()
    plt.clf
    #gc.collect()

def prepare_annexB_summary_table_for_reportlab(df,indi_col="Indicators",total_col="Total(%)",cri_col="Critical_priority(%)",high_col="High_priority(%)",med_col="Medium_priority(%)"):
    df = df.fillna("NA").loc[:,[indi_col,total_col, cri_col,high_col,med_col]]
    lstind = ["Blood culture\ncontamination rate*", "Proportion of notifiable\nantibiotic-pathogen\ncombinations**","Proportion of isolates with\ninfrequent phenotypes or\npotential errors in AST results\n***"]
    for i in range(len(lstind)):
        try:
            df.loc[i,indi_col] = lstind[i]
        except:
            pass    
    #df.at[:,indi_col] = ["Blood culture\ncontamination rate*", "Proportion of notifiable\nantibiotic-pathogen\ncombinations**","Proportion of isolates with\ninfrequent phenotypes or\npotential errors in AST results\n***"]
    for idx in df.index:
        for col in [1,2,3,4]:
            df.iloc[idx,col] = df.iloc[idx,col].replace("(","\n(")
    df = df.rename(columns={total_col:"Total\n(n)",cri_col:"Critical priority\n(n)",high_col:"High priority\n(n)",med_col:"Medium priority\n(n)"})
    df_col = [[indi_col,"Number of observations","","",""],["","Total\n(n)","Critical priority\n(n)","High priority\n(n)","Medium priority\n(n)"]]
    return df_col + df.values.tolist()

def prepare_annexB_summary_table_bymonth_for_reportlab(df, col_month="month", 
                                                       col_rule1="summary_blood_culture_contamination_rate(%)", 
                                                       col_rule2="summary_proportion_of_notifiable_antibiotic-pathogen_combinations(%)", 
                                                       col_rule3="summary_proportion_of_potential_errors_in_the_AST_results(%)"):
    df = df.fillna("NA").loc[:,[col_month, col_rule1, col_rule2, col_rule3]].rename(columns={col_month:"Month", 
                                                                                            col_rule1:"Blood culture\ncontamination rate\n(n)*",
                                                                                            col_rule2:"Proportion of notifiable\nantibiotic-pathogen combinations\n(n)**", 
                                                                                            col_rule3:"Proportion of isolates with\ninfrequent phenotypes or\npotential errors in AST results\n(n)***"})
    return [list(df.columns)] + df.values.tolist()

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
def report_title2(c,title_name,pos_x,pos_y,font_color,font_size=20):
    style = ParagraphStyle('normal',fontName='Helvetica',fontSize=font_size)
    p = Paragraph("<b>" + title_name + "</b>", style)
    #c.setFont("Helvetica-Bold", font_size) # define a large bold Helvetica
    #c.setFillColor(font_color) #define font color
    #c.drawString(pos_x,pos_y,title_name)
    p.wrapOn(c, 500, 20)
    p.drawOn(c, pos_x, pos_y)

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

def report_todaypage(c,pos_x,pos_y,footer_information):
    c.setFont("Helvetica", 9) # define a large bold font
    c.setFillColor('#3e4444')
    c.drawString(pos_x,pos_y,footer_information)

def report1_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('ALIGN',(0,0),(-3,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report2_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'), 
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONT',(0,-1),(0,-1),'Helvetica-Bold'), 
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report3_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'), 
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], 
                 colWidths=[2.2*inch,1.8*inch,1.0*inch,0.9*inch,1.0*inch])

def report6_table(df):
    return Table(df,style=[('FONT',(0,0),(-1,0),'Helvetica-Bold'), 
                           ('FONT',(1,1),(-1,-1),'Helvetica-BoldOblique'),
                           ('FONTSIZE',(0,0),(-1,-1),11),
                           ('TEXTCOLOR',(1,1),(-1,-1),colors.darkblue),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report2_table_nons(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report_table_annexA_page1(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),14),
                           ('GRID',(0,0),(-1,-1),0.5,colors.white),
                           ('ALIGN',(0,0),(-1,-1),'LEFT'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], 
                 colWidths=[3.0*inch,3.0*inch])

def report_table_annexA_page2(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
                           ('ALIGN',(0,-1),(0,-1),'LEFT')], 
                 colWidths=[1.3*inch,0.9*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch,0.5*inch])

def report_table_annexA_page3(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,0),9),
                           ('FONTSIZE',(0,1),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')], 
                 colWidths=[1.2*inch,1.0*inch,1.0*inch])

def report_table_annexB(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica-Bold'),
                           ('FONT',(0,1),(-1,-1),'Helvetica'),
                           ('FONTSIZE',(0,0),(-1,-1),9),
                           ('GRID',(0,0),(-1,-1),0.5,colors.darkgrey),
                           ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                           ('ALIGN',(0,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE')])

def report_table_annexB_page1(df):
    return Table(df,style=[('FONT',(0,0),(-1,-1),'Helvetica'),
                           ('FONT',(0,0),(-1,1),'Helvetica-Bold'),
                           ('FONTSIZE',(0,0),(-1,-1),10),
                           ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                           ('ALIGN',(0,2),(0,-1),'LEFT'),
                           ('ALIGN',(0,0),(0,1),'CENTER'),
                           ('ALIGN',(1,0),(-1,-1),'CENTER'),
                           ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
                           ('SPAN',(0,0),(0,1)), 
                           ('SPAN',(1,0),(-1,0))])
def canvas_printpage_nototalpage(c,curpage,today=date.today().strftime("%d %b %Y"),bisrotate90=False,ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme=""):
    report_todaypage(c,55,30,"Created on: "+today)
    imode = ipagemode
    if imode == 2:
        report_todaypage(c,250,30,ssectionanme + " Page " + str(curpage))
    else:
        report_todaypage(c,270,30,"Page " + str(curpage))
    if bisrotate90 == True:
        c.rotate(90)
    c.showPage()
def canvas_printpage(c,curpage,lastpage,today=date.today().strftime("%d %b %Y"),bisrotate90=False,ipagemode=AC.CONST_REPORTPAGENUM_MODE,ssectionanme="",isecmaxpage=1,istartpage=1):
    report_todaypage(c,55,30,"Created on: "+today)
    imode = ipagemode
    if imode == 1:
        report_todaypage(c,270,30,"Page " + str(curpage) + " of " + str(lastpage))
    elif imode == 2:
        report_todaypage(c,250,30,ssectionanme + " Page " + str(curpage) + " of " + str(isecmaxpage))
    else:
        iseccurpage = curpage - istartpage + 1
        if iseccurpage <= isecmaxpage:
            report_todaypage(c,270,30,"Page " + str(curpage) + " of " + str(lastpage))
        else:
            iasc = 65 + iseccurpage - isecmaxpage - 1
            curpage = istartpage + isecmaxpage -1
            report_todaypage(c,270,30,"Page " + str(curpage) + "-" + str(chr(iasc)) +" of " + str(lastpage))
    if bisrotate90 == True:
        c.rotate(90)
    c.showPage()