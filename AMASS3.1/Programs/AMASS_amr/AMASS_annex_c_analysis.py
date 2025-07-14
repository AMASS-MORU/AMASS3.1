
#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: CHALIDA RAMGSIWUTISAK
# Created on: 01 SEP 2023
# Last update on: 20 JUNE 2025 #v3.1 3102
import pandas as pd
import psutil,gc
import numpy as np #v3.1 3102
import datetime #for setting date-time format
import logging #for creating logfile
from reportlab.lib.pagesizes import A4 #for setting PDF size
from reportlab.pdfgen import canvas #for creating PDF page
from reportlab.platypus.paragraph import Paragraph #for creating text in paragraph
from reportlab.lib.styles import ParagraphStyle #for setting paragraph style
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY #for setting paragraph style
from reportlab.platypus import * #for plotting graph and tables
from reportlab.lib.colors import * #for importing color palette
from reportlab.graphics.shapes import Drawing #for creating shapes
from reportlab.lib.units import inch #for importing inch for plotting
from reportlab.lib import colors #for importing color palette
from reportlab.platypus.flowables import Flowable #for plotting graph and tables
import AMASS_amr_const as AC
import AMASS_annex_c_const as ACC
import AMASS_amr_commonlib as AL

#######################################################################
#######################################################################
#####Step1. Merged Hospital-Microbiology data to SaTScan's input#######
#######################################################################
#######################################################################
def prepare_fromHospMicro_toSaTScan(logger,df_all=pd.DataFrame(),df_blo=pd.DataFrame()):
    AL.printlog("Start Cluster signal identification (ANNEX C): " + str(datetime.datetime.now()),False,logger)
    #Preparing configurations from Configuration.xlsx
    df_config = AL.readxlsxorcsv(AC.CONST_PATH_ROOT + "Configuration/", "Configuration",logger)
    df_config_satscan = df_config.loc[df_config[ACC.CONST_COL_AMASS_PRMNAME].str.contains("satscan_")]
    dict_configuration_prm_final = prepare_configuration_to_constant(df=df_config_satscan,dict_const=ACC.dict_configuration_prm,dict_intermediate=ACC.dict_intermediate_configtosatscan,
                                                                    col_amassprm=ACC.CONST_COL_AMASS_PRMNAME,col_userprm=ACC.CONST_COL_USER_PRMNAME,prefix="satscan_")
    df_config_profile = df_config.loc[df_config[ACC.CONST_COL_AMASS_PRMNAME].str.contains("profiling_")]
    dict_configuration_profile_final = prepare_configuration_to_constant(df=df_config_profile,dict_const=ACC.dict_configuration_profile,
                                                                        col_amassprm=ACC.CONST_COL_AMASS_PRMNAME,col_userprm=ACC.CONST_COL_USER_PRMNAME,prefix="profiling_")
    for lo_org in ACC.dict_ast.keys():
    # for lo_org in ["organism_staphylococcus_aureus"]:
        lst_usedcolumns = []
        if lo_org in ACC.dict_configuration_astforprofile.keys():
            # lst_usedcolumns = [AC.CONST_NEWVARNAME_ORG3]+AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)[lo_org][4]+ACC.dict_configuration_astforprofile[lo_org]
            lst_usedcolumns = [AC.CONST_NEWVARNAME_ORG3]+AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)[lo_org][4]+ACC.dict_configuration_astforprofile[lo_org][0] #v3.1 3102
        else:
            lst_usedcolumns = [AC.CONST_NEWVARNAME_ORG3]+AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)[lo_org][4]

        df_blo_str = df_blo.copy()
        df_all_str = df_all.copy()
        #setting values of these columns from categories to string or int
        df_blo_str = assign_strtypetocolumns(df=df_blo_str, lst_col=lst_usedcolumns)
        df_all_str = assign_strtypetocolumns(df=df_all_str, lst_col=lst_usedcolumns)
                
        lst_ris_rpt2=[]
        lst_ris_rpt2_export=[]
        try:
            #getting list of antibiotics for each pathogen i.e. S. aureus 
            lst_ris_rpt2 = get_lstastforpathogen(lo_org=lo_org,check_writereport=False)
            lst_ris_rpt2_export = get_lstastforpathogen(lo_org=lo_org,check_writereport=True)
        except Exception as e:
            AL.printlog("Error, ANNEX C antibiotic selection for profiling: " +  str(e),True,logger)
        for lst_value in ACC.dict_ast[lo_org]:
            #selecting organisms; A. baumannii
            df_blo_sp = select_dfbyOrganism(logger,df=df_blo_str, col_org=AC.CONST_NEWVARNAME_ORG3, str_selectorg=lo_org) 
            df_all_sp = select_dfbyOrganism(logger,df=df_all_str, col_org=AC.CONST_NEWVARNAME_ORG3, str_selectorg=lo_org)
            #selecting AMR pathogen; CRAB
            df_blo_sp_amr = select_resistantProfile(df=df_blo_sp, d_ast_val=lst_value)
            df_all_sp_amr = select_resistantProfile(df=df_all_sp, d_ast_val=lst_value)
            df_blo_sp_amr[ACC.CONST_COL_AMRPATHOGEN] = lst_value[-1] #named CRAB
            df_all_sp_amr[ACC.CONST_COL_AMRPATHOGEN] = lst_value[-1] #named CRAB
            df_blo_sp_amr.to_excel(AC.CONST_PATH_TEMPWITH_PID+"df_blo_sp_amr_"+lst_value[-1]+".xlsx")
            df_all_sp_amr.to_excel(AC.CONST_PATH_TEMPWITH_PID+"df_all_sp_amr_"+lst_value[-1]+".xlsx")
            #deduplicating data
            df_blo_sp_amr_dedup = fn_deduplicateannexc_hospmico(df_blo_sp_amr,ACC.CONST_COL_AMRPATHOGEN,lst_value[-1])
            df_all_sp_amr_dedup = fn_deduplicateannexc_hospmico(df_all_sp_amr,ACC.CONST_COL_AMRPATHOGEN,lst_value[-1])
            df_blo_sp_amr_dedup.to_excel(AC.CONST_PATH_TEMPWITH_PID+"df_blo_sp_amr_dedup_"+lst_value[-1]+".xlsx")
            df_all_sp_amr_dedup.to_excel(AC.CONST_PATH_TEMPWITH_PID+"df_all_sp_amr_dedup_"+lst_value[-1]+".xlsx")
            #selecting HO for AnnexC
            df_blo_sp_amr_dedup_ho = df_blo_sp_amr_dedup.loc[df_blo_sp_amr_dedup[AC.CONST_NEWVARNAME_COHO_FINAL]==1,:]
            df_all_sp_amr_dedup_ho = df_all_sp_amr_dedup.loc[df_all_sp_amr_dedup[AC.CONST_NEWVARNAME_COHO_FINAL]==1,:]
            print ("-----------------------")
            print ("-----------------------")
            print ("-----------------------")
            AL.printlog("AMR pathogen : "   + lst_value[-1],False,logger)
            AL.printlog("Specimen model : Blood",False,logger)
            AL.printlog("No. patients with positive isolates for "+lo_org+" : "+str(len(set(df_blo_sp[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" : "+str(len(set(df_blo_sp_amr[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" (deduplicated) : "+str(len(set(df_blo_sp_amr_dedup[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" and admitted for >2 calendar days : "+str(len(set(df_blo_sp_amr_dedup_ho[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("Specimen model : All specimens",False,logger)
            AL.printlog("No. patients with positive isolates for "+lo_org+" : "+str(len(set(df_all_sp[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" : "+str(len(set(df_all_sp_amr[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" (deduplicated) : "+str(len(set(df_all_sp_amr_dedup[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            AL.printlog("No. patients with positive isolates for "+lst_value[-1]+" and admitted for >2 calendar days : "+str(len(set(df_all_sp_amr_dedup_ho[AC.CONST_VARNAME_HOSPITALNUMBER]))),False,logger)
            #selecting profiles from configuration
            #do antibiotics selection for profiling
            lst_ris_profiletemp_blo= select_atbforprofiling(logger, df=df_blo_sp_amr_dedup_ho, lst_col_ris=lst_ris_rpt2, configuration_profile=dict_configuration_profile_final) #v3.1 3102
            lst_ris_profile_blo    = lst_ris_profiletemp_blo[0] #v3.1 3102
            lst_ris_profile_blo_c1 = lst_ris_profiletemp_blo[1] #v3.1 3102
            lst_ris_profile_blo_c2 = lst_ris_profiletemp_blo[2] #v3.1 3102
            # print (lst_ris_profile_blo,lst_ris_profile_blo_c1,lst_ris_profile_blo_c2)
            profiletemp = ",".join(lst_ris_profile_blo).replace("RIS","") #v3.1 3102
            AL.printlog("Antibiotics for profiling from blood :" + profiletemp,False,logger) #v3.1 3102
            #do antibiotics selection for profiling by using isolates positive to that pathogen in clinical specimen
            lst_ris_profiletemp_all= select_atbforprofiling(logger, df=df_all_sp_amr_dedup_ho, lst_col_ris=lst_ris_rpt2, configuration_profile=dict_configuration_profile_final) #v3.1 3102
            lst_ris_profile_all    = lst_ris_profiletemp_all[0] #v3.1 3102
            lst_ris_profile_all_c1 = lst_ris_profiletemp_all[1] #v3.1 3102
            lst_ris_profile_all_c2 = lst_ris_profiletemp_all[2] #v3.1 3102
            # print (lst_ris_profile_all,lst_ris_profile_all_c1,lst_ris_profile_all_c2)
            profiletemp = ",".join(lst_ris_profile_all).replace("RIS","") #v3.1 3102
            AL.printlog("Antibiotics for profiling from clinical specimens :" + profiletemp,False,logger) #v3.1 3102
            #profiling
            df_blo_sp_amr_dedup_ho_profile = pd.DataFrame()
            df_all_sp_amr_dedup_ho_profile = pd.DataFrame()
            try:
                df_blo_sp_amr_dedup_ho_profile  = create_risprofile(logger,df=df_blo_sp_amr_dedup_ho, lst_col_ris=lst_ris_rpt2, lst_col_ristemp=lst_ris_profile_blo, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP) #v3.1 3102
                df_all_sp_amr_dedup_ho_profile  = create_risprofile(logger,df=df_all_sp_amr_dedup_ho, lst_col_ris=lst_ris_rpt2, lst_col_ristemp=lst_ris_profile_all, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP) #v3.1 3102
            except Exception as e:
                AL.printlog("Error, ANNEX C profiling: " +  str(e),True,logger)
            #creating dictionary for profile
            d_profile    = create_dictforMapProfileID_byorg(logger, df_all=df_all_sp_amr_dedup_ho, df_blo=df_blo_sp_amr_dedup_ho, sh_org=lst_value[-1])
            #exporting d_profile to profile_information.xlsx
            export_dprofile(logger, df_groupprofile=d_profile[0], lst_col_ris=lst_ris_rpt2, lst_col_rpt2_export=lst_ris_rpt2_export, filename_profile=AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PROFILE+"_"+lst_value[-1].upper()) #v3.1 3102
            #exporting sumamtion of AST results to ast_information.xlsx
            df_blo_sp_amr_dedup_ho_ris = summation_astresult(logger, df=df_blo_sp_amr_dedup_ho, dict_configuration=dict_configuration_profile_final, lst_col_ris=lst_ris_rpt2, lst_col_rpt2_export=lst_ris_rpt2_export, lst_atb_c1=lst_ris_profile_blo_c1, lst_atb_c2=lst_ris_profile_blo_c2, lst_atb_c1_c2=lst_ris_profile_blo, col_spctype=AC.CONST_NEWVARNAME_AMASSSPECTYPE, sh_spc="blo") #v3.1 3102
            df_all_sp_amr_dedup_ho_ris = summation_astresult(logger, df=df_all_sp_amr_dedup_ho, dict_configuration=dict_configuration_profile_final, lst_col_ris=lst_ris_rpt2, lst_col_rpt2_export=lst_ris_rpt2_export, lst_atb_c1=lst_ris_profile_all_c1, lst_atb_c2=lst_ris_profile_all_c2, lst_atb_c1_c2=lst_ris_profile_all, col_spctype=AC.CONST_NEWVARNAME_AMASSSPECTYPE, sh_spc="all") #v3.1 3102
            export_astresult(logger, lst_df=[df_blo_sp_amr_dedup_ho_ris,df_all_sp_amr_dedup_ho_ris], filename_ast=AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_AST+"_"+lst_value[-1].upper()) #v3.1 3102
            #mapping profile to dataframe -- blood
            if len(df_blo_sp_amr_dedup_ho_profile)>0:
                df_blo_sp_amr_dedup_ho_profile = map_profileIDtoDataframe(logger, df=df_blo_sp_amr_dedup_ho_profile, d_profile=d_profile[1], sh_org=lst_value[-1], sh_spc="blo", col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID)
                try:
                    df_blo_sp_amr_dedup_ho_profile.drop(columns=[ACC.CONST_COL_PROFILE,ACC.CONST_COL_PROFILETEMP]).to_excel(AC.CONST_PATH_REPORTWITH_PID+ACC.CONST_FILENAME_HO_DEDUP+"_"+lst_value[-1].upper()+"_"+ACC.dict_spc["blo"]+".xlsx", index=False)
                except Exception as e:
                    AL.printlog("Error, ANNEX C exporting "+ACC.CONST_FILENAME_HO_DEDUP+"_"+lst_value[-1].upper()+"_"+ACC.dict_spc["blo"]+".xlsx"+": " +  str(e),True,logger)
            #mapping profile to dataframe -- all specimens
            if len(df_all_sp_amr_dedup_ho_profile)>0:
                df_all_sp_amr_dedup_ho_profile = map_profileIDtoDataframe(logger, df=df_all_sp_amr_dedup_ho_profile, d_profile=d_profile[2], sh_org=lst_value[-1], sh_spc="all", col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID) #v3.1 3102
                try:
                    df_all_sp_amr_dedup_ho_profile.drop(columns=[ACC.CONST_COL_PROFILE,ACC.CONST_COL_PROFILETEMP]).to_excel(AC.CONST_PATH_REPORTWITH_PID+ACC.CONST_FILENAME_HO_DEDUP+"_"+lst_value[-1].upper()+"_"+ACC.dict_spc["all"]+".xlsx", index=False)
                except Exception as e:
                    AL.printlog("Error, ANNEX C exporting "+ACC.CONST_FILENAME_HO_DEDUP+"_"+lst_value[-1].upper()+"_"+ACC.dict_spc["all"]+".xlsx"+": " +  str(e),True,logger)
            #preparing SaTScan's inputs
            evaluation_study = retrieve_startEndDate(filename=AC.CONST_PATH_RESULT+AC.CONST_FILENAME_sec1_res_i) #for satscan_param.prm
            for sh_spc in ACC.dict_spc.keys():
                prepare_satscanInput(logger, 
                                    configuration_prm=dict_configuration_prm_final,
                                    filename_isolate =ACC.CONST_FILENAME_HO_DEDUP, 
                                    filename_ward    =ACC.CONST_FILENAME_WARD, 
                                    filename_case    =ACC.CONST_FILENAME_INPUT, 
                                    filename_oriparam=ACC.CONST_FILENAME_ORIPARAM, 
                                    filename_newparam=ACC.CONST_FILENAME_NEWPARAM, 
                                    path_input=AC.CONST_PATH_REPORTWITH_PID,
                                    path_output=AC.CONST_PATH_TEMPWITH_PID,
                                    sh_org=lst_value[-1], 
                                    sh_spc=sh_spc, 
                                    start_date=evaluation_study[0], 
                                    end_date  =evaluation_study[1])
        del [[df_blo_sp, df_blo_sp_amr, df_blo_sp_amr_dedup, df_blo_sp_amr_dedup_ho, df_blo_sp_amr_dedup_ho_profile]]
        del [[df_all_sp, df_all_sp_amr, df_all_sp_amr_dedup, df_all_sp_amr_dedup_ho, df_all_sp_amr_dedup_ho_profile]]
        gc.collect()
        df_blo_sp=pd.DataFrame()
        df_blo_sp_amr=pd.DataFrame()
        df_blo_sp_amr_dedup=pd.DataFrame()
        df_blo_sp_amr_dedup_ho=pd.DataFrame()
        df_blo_sp_amr_dedup_ho_profile=pd.DataFrame()
        df_all_sp=pd.DataFrame()
        df_all_sp_amr=pd.DataFrame()
        df_all_sp_amr_dedup=pd.DataFrame()
        df_all_sp_amr_dedup_ho=pd.DataFrame()
        df_all_sp_amr_dedup_ho_profile=pd.DataFrame()
    del [[df_blo, df_blo_str]]
    del [[df_all, df_all_str]]
    gc.collect()
    df_blo=pd.DataFrame()
    df_blo_str=pd.DataFrame()
    df_all=pd.DataFrame()
    df_all_str=pd.DataFrame()

#####Small functions in the step1.1#####
#Mapping configs from Configuration file to annex_c_const
def prepare_configuration_to_constant(df=pd.DataFrame(), dict_const={}, dict_intermediate={}, prefix="", col_amassprm="", col_userprm=""):
    dict_const_mapped = dict_const
    for amass_prm in df.loc[:,col_amassprm]:
        user_prm = df.loc[df[col_amassprm]==amass_prm,col_userprm].tolist()[0]
        satscan_value = ""
        try:
            satscan_value = dict_intermediate[amass_prm][user_prm]
        except:
            if (isinstance(user_prm, int)) or (isinstance(user_prm, float)):
                satscan_value = user_prm
            else:
                print ("ERROR: Please check configuarion file.")

        for keys,values in dict_const.items():
            if amass_prm == values:
                if prefix == "satscan_":
                    dict_const_mapped[keys] = str(satscan_value)
                elif prefix == "profiling_":
                    dict_const_mapped[keys] = satscan_value
    return dict_const_mapped

#Assigning dtype as string for all used columns
#Original dtype is category
def assign_strtypetocolumns(df=pd.DataFrame(), lst_col=[]):
    lst_col_new = []
    for col in lst_col:
        if (col not in lst_col_new) and (col in df.columns):
            lst_col_new.append(col)
    try:
        df[lst_col_new] = df[lst_col_new].astype(str).fillna("").replace("","-")
    except:
        for col in lst_col_new:
            try:
                df[col] = df[col].astype(str).fillna("").replace("","-")
            except:
                pass
    return df
#Parsing df from core AMASS to the step before grouping isolates by specimen date.
#Resistant profile, deduplication
#input : filename of HospMicroData
#output: dataframe of deduplicated hospital-origin isolates with profiles
def fn_deduplicatedata(df,list_sort,list_order,na_posmode,list_dupcolchk,keepmode) :
    return df.sort_values(by = list_sort, ascending = list_order, na_position = na_posmode).drop_duplicates(subset=list_dupcolchk, keep=keepmode)
# Filter orgcat before dedup (For merge data) with admission date
def fn_deduplicateannexc_hospmico(df,colname,orgcat) :
    #return fn_deduplicatedata(df,[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AMR,AC.CONST_NEWVARNAME_AMR_TESTED,AC.CONST_NEWVARNAME_CLEANADMDATE],[True,True,False,False,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
    #Change in 3.1 3031
    return fn_deduplicatedata(df.loc[df[colname]==orgcat],
                              [AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AST_R,AC.CONST_NEWVARNAME_AST_I,AC.CONST_NEWVARNAME_AST_TESTED,AC.CONST_NEWVARNAME_CLEANADMDATE],
                              [True,True,False,False,False,True],
                              "last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
def summation_numpatient(logger,df=pd.DataFrame(),lst_col=[],col_forgrouping=""):
    df_ = df
    try:
        df_[col_forgrouping] = 1
        df_ = df_.loc[:,lst_col+[col_forgrouping]].groupby([col for col in lst_col if col != col_forgrouping]).sum()
    except Exception as e:
        AL.printlog("Error, ANNEX C counting number of patient: " + str(e),True,logger)
    return df_

#Summarising AST results of each antibiotic for Supplementary report AnnexC
#input : df before SaTScan input
#output: df with index are ["R","I","S","Total"] and columns are lst of antibiotics + [col_numprofile]
#v3.1 3102
def summation_astresult(logger, df=pd.DataFrame(), dict_configuration={}, lst_col_ris=[], lst_col_rpt2_export=[], lst_atb_c1=[], lst_atb_c2=[], lst_atb_c1_c2=[], col_spctype="", sh_spc=""):
    df_temp = df.copy()
    #criteria for %tested isolates
    str_perctested = int(select_configvalue(configuration_user=dict_configuration,configuration_default=ACC.dict_configuration_profile_default,b_satscan=False,param=ACC.CONST_VALUE_MIN_TESTATBRATE))
    lst_ris = ACC.lst_ris #v3.1 3102
    lst_row = lst_ris+[ACC.CONST_VALUE_NA,ACC.CONST_VALUE_C1_1+str(str_perctested)+ACC.CONST_VALUE_C1_2,ACC.CONST_VALUE_C2,ACC.CONST_VALUE_SUMMARY] #v3.1 3102
    # df_ris    = pd.DataFrame(index=lst_ris)
    df_ris_sum= pd.DataFrame(index=lst_row)
    df_merge  = pd.DataFrame()
    if len(df_temp)>0:
        df_temp  = df_temp[lst_col_ris].fillna("NA").replace("nan","NA").replace("-","NA").apply(lambda col: col.value_counts()).fillna(0).astype(int) # R-AMK: 6, I-AMK: 7, S-AMK: 5, NA: 2 #v3.1 3102
        df_merge = pd.concat([df_ris_sum,df_temp],axis=1,ignore_index=False,sort=False).fillna(0)
        #assigning "-" to any columns that have not test AST results
        lst_atb_passc   = list(set(lst_atb_c1_c2 + lst_atb_c1 + lst_atb_c2))
        lst_atb_notpassc= list(set(lst_col_ris) - set(lst_atb_passc))
        df_merge[lst_atb_notpassc] = df_merge[lst_atb_notpassc].astype(object)
        df_merge.loc[:,lst_atb_notpassc] = ACC.CONST_VALUE_NOTTESTEDATB
        #assigning "P" to any columns that passed each criteria
        df_merge.loc[lst_row[-3],lst_atb_c1]   = ACC.CONST_VALUE_PASSEDATB
        df_merge.loc[lst_row[-2],lst_atb_c2]   = ACC.CONST_VALUE_PASSEDATB
        df_merge.loc[lst_row[-1],lst_atb_c1_c2]= ACC.CONST_VALUE_PASSEDATB
        #assigning "F" to any columns that passed each criteria
        df_merge.loc[lst_row[-3],lst_atb_passc] = df_merge.loc[lst_row[-3],lst_atb_passc].replace(0,ACC.CONST_VALUE_NOTPASSEDATB)
        df_merge.loc[lst_row[-2],lst_atb_passc] = df_merge.loc[lst_row[-2],lst_atb_passc].replace(0,ACC.CONST_VALUE_NOTPASSEDATB)
        df_merge.loc[lst_row[-1],lst_atb_passc] = df_merge.loc[lst_row[-1],lst_atb_passc].replace(0,ACC.CONST_VALUE_NOTPASSEDATB)
        df_merge.columns = lst_col_rpt2_export
        df_merge[col_spctype] = sh_spc
    else:
        pass
    return df_merge

#Concatinating dataframes of AST results for Supplementary report AnnexC
#v3.1 3102
def export_astresult(logger, lst_df=[], filename_ast="", col_profileid=ACC.CONST_COL_PROFILEID):
    try:
        df_concat = pd.concat(lst_df,axis=0,ignore_index=False,sort=False).fillna(0).reset_index().rename(columns={"index":col_profileid})
        if len(df_concat)>0:
            df_concat.to_excel(filename_ast+".xlsx", index=False)
        else:
            pass
    except Exception as e:
        AL.printlog("Error, ANNEX C merging dataframes for ast results: " + str(e),True,logger)

#Creating dictionary for mappingID
#ID ordering by all specimen and then blood of each organism
#d_profile["profile_sequence"]="profile_ID"
#v3.1 3102
def create_dictforMapProfileID_byorg(logger, df_all=pd.DataFrame(), df_blo=pd.DataFrame(), sh_org="", col_profileid=ACC.CONST_COL_PROFILEID, col_profile=ACC.CONST_COL_PROFILE, col_profiletemp=ACC.CONST_COL_PROFILETEMP, col_numprofile_all=ACC.CONST_COL_NUMPROFILE_ALL, col_numprofile_blo=ACC.CONST_COL_NUMPROFILE_BLO):
    prefixID_blo=sh_org.upper()+"_BL_" #v3.1 3102
    prefixID_all=sh_org.upper()+"_ALL_"
    count_blo = 0
    count_all = 0
    d_profile_blo = {}
    d_profile_all = {}
    df_blo_num = summation_numpatient(logger,df_blo,lst_col=[col_profile,col_profiletemp],col_forgrouping=col_numprofile_blo).reset_index()
    df_all_num = summation_numpatient(logger,df_all,lst_col=[col_profile,col_profiletemp],col_forgrouping=col_numprofile_all).reset_index()
    #blood
    for idx in df_blo_num.index:
        try:
            if df_blo_num.loc[idx,col_profiletemp] == "":
                df_blo_num.at[idx,col_profileid] = prefixID_blo+"na"
            else:
                count_blo += 1
                df_blo_num.at[idx,col_profileid] = prefixID_blo+str(count_blo)
            d_profile_blo[df_blo_num.loc[idx,col_profile]] = df_blo_num.loc[idx,col_profileid]
        except Exception as e:
            AL.printlog("Error, ANNEX C assigning profile ID: " + str(e),True,logger)
    #all
    for idx in df_all_num.index:
        try:
            if df_all_num.loc[idx,col_profiletemp] == "":
                df_all_num.at[idx,col_profileid] = prefixID_all+"na"
            else:
                count_all += 1
                df_all_num.at[idx,col_profileid] = prefixID_all+str(count_all)
            d_profile_all[df_all_num.loc[idx,col_profile]] = df_all_num.loc[idx,col_profileid]
        except Exception as e:
            AL.printlog("Error, ANNEX C assigning profile ID: " + str(e),True,logger)
    #merging
    df = pd.concat([df_blo_num,df_all_num],axis=0).fillna(0).reset_index()
    if len(df)>0:
        df = df.loc[:,[col_profileid,col_profile,col_numprofile_blo,col_numprofile_all]]
    else:
        df = pd.DataFrame(columns=[col_profileid,col_profile,col_numprofile_blo,col_numprofile_all])
    return df, d_profile_blo, d_profile_all

def map_profileIDtoDataframe(logger, df=pd.DataFrame(), d_profile={}, filename_dedup="", sh_org="", sh_spc="", 
                             col_profile=ACC.CONST_COL_PROFILE, col_profileid=ACC.CONST_COL_PROFILEID):
    try:
        df[col_profileid] = df[col_profile].map(d_profile).fillna("")
        return df
    except Exception as e:
        AL.printlog("Error, ANNEX C mapping profile id: " + str(e),True,logger)
#Creating profile information.xlsx
#column profile_ID : profile_CRAB_1, profile_CRAB_2, profile_CRAB_3, ..., profile_XXXX_N
#column RISdrug1 : -, -, R, S
#column RISdrug2 : R, -, -, -
#column RISdrugN : -, -, R, S
#v3.1 3102
def export_dprofile(logger, df_groupprofile=pd.DataFrame(), lst_col_ris=[], lst_col_rpt2_export=[], filename_profile="", 
                    col_profileid=ACC.CONST_COL_PROFILEID, col_profile=ACC.CONST_COL_PROFILE, col_numprofile_all=ACC.CONST_COL_NUMPROFILE_ALL, col_numprofile_blo=ACC.CONST_COL_NUMPROFILE_BLO):
    if len(df_groupprofile) > 0:
        try:
            df_profile = df_groupprofile.copy()
            df_profile[lst_col_rpt2_export] = df_profile[col_profile].str.split(";",expand=True)
            df_profile = df_profile.loc[:,[col_profileid]+lst_col_rpt2_export+[col_numprofile_blo,col_numprofile_all]]
            df_profile.to_excel(filename_profile+".xlsx", index=False)
        except Exception as e:
            AL.printlog("Error, ANNEX C exporting profile_information.xlsx: " + str(e),True,logger)
    else:
        pass
#selecting organisms
#v3.1 3102
def select_dfbyOrganism(logger, df=pd.DataFrame(), col_org="", str_selectorg=""):
    df_new = pd.DataFrame()
    try:
        df_new = df.loc[(df[col_org]==str_selectorg)]
    except Exception as e:
        AL.printlog("Error, ANNEX C Selecting organism: " +  str(e),True,logger)
    return df_new
#hospitalized selection
def select_ho(logger, df=pd.DataFrame(), col_origin="", num_ho=0):
    df_ho = pd.DataFrame()
    try:
        df_ho = df.loc[(df[col_origin]==num_ho)]
    except Exception as e:
        AL.printlog("Error, ANNEX C Hospital-origin isolate selection: " +  str(e),True,logger)
    return df_ho
#resistant isolate selection
def select_resistantProfile(df=pd.DataFrame(), d_ast_val=[]):
    df_ast = pd.DataFrame()
    if len(d_ast_val) == 3:
        df_ast = df.loc[(df[d_ast_val[0]]==str(d_ast_val[1]))]
    elif len(d_ast_val) > 3:
        i = 0
        df_ast = df.copy()
        while (i < len(d_ast_val)) and (i+1!=len(d_ast_val)): #in the case; ["RIS3gc","R","RISCarbapenem","S","3gcr-csec"] >>> do not retrieve "3gcr-crec"
            df_ast = df_ast.loc[(df[d_ast_val[i]]==str(d_ast_val[i+1]))]
            i+=2
    return df_ast
#Getting list of antibiotics from report2 i.e. S. aureus is ['RISVancomycin', 'RISClindamycin', 'RISCefoxitin', 'RISOxacillin']
def get_lstastforpathogen(lo_org="",check_writereport=False):
    lst_atb      = []
    lst_atbgroup = []
    #getting list ['RIS3gc','RISCarbapenem','RISFluoroquin','RISTetra','RISaminogly','RISmrsa','RISpengroup']
    for val in AC.dict_atbgroup().values():
        lst_atbgroup.append(val[0])
    #getting list of antibiotics for analysis(column name) or report except lst_atbgroup
    d_ = AC.get_dict_orgcatwithatb(bisabom=True,bisentspp=True)
    v = [i for i in range(len(d_[lo_org][4])) if d_[lo_org][4][i] not in lst_atbgroup]
    if check_writereport:
        lst_atb = [d_[lo_org][3][idx].replace("RIS","") for idx in v ]
    else:
        lst_atb = [d_[lo_org][4][idx] for idx in v ]
    #getting list of additional antibiotics from configuration
    if (ACC.dict_configuration_astforprofile != {}) and (lo_org in ACC.dict_configuration_astforprofile.keys()):
        # lst_atb = lst_atb + ACC.dict_configuration_astforprofile[lo_org]
        lst_atb = lst_atb + ACC.dict_configuration_astforprofile[lo_org][0] #v3.1 3102
        lst_atb_unique = []
        for atb in lst_atb:
            if atb not in lst_atb_unique:
                if check_writereport:
                    lst_atb_unique.append(atb.replace("RIS","").replace("Piperacillin_and_tazobactam","Piperacillin/tazobactam").replace("Sulfamethoxazole_and_trimethoprim","Co-trimoxazole"))
                else:
                    lst_atb_unique.append(atb)
            else:
                pass
    else:
        lst_atb_unique = lst_atb
    return lst_atb_unique
#Selecting range of criteria for selecting antibiotics profiling
#In the case that there is no setting >>> using default setting
def select_configvalue(configuration_user={},configuration_default={},b_satscan=True,param=""):
    val_ = ""
    if b_satscan:
        if "satscan_" in str(configuration_user[param]):
            val_ = str(configuration_default[param]) #real user value didn't mapped and disappear
        else:
            val_ = str(configuration_user[param]) #real user value is mapped completely
    else:
        try:
            val_ = float(configuration_user[param])
        except:
            val_ = float(configuration_default[param])
    return val_
#Counting number of resistant, intermediate, and susceptible
#In the case that there is no setting >>> assign 0
def count_numRISbyATB(df=pd.DataFrame(),col_atb=[],val_forcount=""):
    num = 0
    try:
        num = int(df[col_atb].str.count(val_forcount).sum())
    except: #no val_forcount in a column >>> num = 0
        pass
    return num
#Retrieving minimum tested ATB and maximum tested ATB i.e. minimum is 0.1%, maximum is 99.9%
def retrieve_minMaxTestedATB(dict_config_user={}, dict_config_default={},param_min="",param_max=""):
    min_tested = 0
    max_tested = 0
    if param_min in dict_config_default.keys():
        min_tested = select_configvalue(configuration_user=dict_config_user,configuration_default=dict_config_default,b_satscan=False,param=param_min)
    else:
        pass
    if param_max in dict_config_default.keys():
        max_tested = select_configvalue(configuration_user=dict_config_user,configuration_default=dict_config_default,b_satscan=False,param=param_max)
    else:
        pass
    return min_tested,max_tested
#Selecting an ATB which pass the criteria of tested ATB i.e. 0.1%=<resistant=<99.9%
def select_atb_byminMaxTestedATB(str_atb="",numerator=0,denominator=0,min_tested=0,max_tested=0):
    str_atb_pass = ""
    if (numerator*100/denominator>=min_tested) and (numerator*100/denominator<=max_tested):
        str_atb_pass = str_atb
    else:
        pass
    return str_atb_pass

#Selecting list of ATBs which pass the criteria for profiling
def select_atbforprofiling(logger,df=pd.DataFrame(), lst_col_ris=[], configuration_profile={},
                          resistant=ACC.dict_ris["resistant"],intermediate=ACC.dict_ris["intermediate"],susceptible=ACC.dict_ris["susceptible"]):
    lst_col_ris_include = [] #v3.1 3102
    lst_col_ris_passC1  = [] #v3.1 3102
    lst_col_ris_passC2  = [] #v3.1 3102
    for atb in lst_col_ris:
        try:
            if atb in df.columns:
                num_r = count_numRISbyATB(df=df,col_atb=atb,val_forcount=resistant)
                num_i = count_numRISbyATB(df=df,col_atb=atb,val_forcount=intermediate)
                num_s = count_numRISbyATB(df=df,col_atb=atb,val_forcount=susceptible)
                total_testedatb = num_r+num_i+num_s
                min_testedatb = select_configvalue(configuration_user=configuration_profile,configuration_default=ACC.dict_configuration_profile_default,b_satscan=False,param=ACC.CONST_VALUE_MIN_TESTATBRATE)
                max_testedatb = select_configvalue(configuration_user=configuration_profile,configuration_default=ACC.dict_configuration_profile_default,b_satscan=False,param=ACC.CONST_VALUE_MAX_TESTATBRATE)
                if (total_testedatb*100/len(df)>=min_testedatb) and (total_testedatb*100/len(df)<=max_testedatb):
                    #list of antibiotics passing tested isolates
                    lst_col_ris_passC1.append(atb) #v3.1 3102
                    #calculating RIS rate
                    min_tested_r,max_tested_r = retrieve_minMaxTestedATB(dict_config_user=configuration_profile, dict_config_default=ACC.dict_configuration_profile_default,
                                                                        param_min=ACC.CONST_VALUE_MIN_RRATE, param_max=ACC.CONST_VALUE_MAX_RRATE) #v3.1 3102
                    min_tested_s,max_tested_s = retrieve_minMaxTestedATB(dict_config_user=configuration_profile, dict_config_default=ACC.dict_configuration_profile_default,
                                                                        param_min=ACC.CONST_VALUE_MIN_SRATE, param_max=ACC.CONST_VALUE_MAX_SRATE) #v3.1 3102
                    lst_col_ris_include.append(select_atb_byminMaxTestedATB(str_atb=atb,numerator=num_r,denominator=total_testedatb,min_tested=min_tested_r,max_tested=max_tested_r)) #v3.1 3102
                    lst_col_ris_include.append(select_atb_byminMaxTestedATB(str_atb=atb,numerator=num_s,denominator=total_testedatb,min_tested=min_tested_s,max_tested=max_tested_s)) #v3.1 3102
                    print (atb, "%tested AST:"+str(round(total_testedatb*100/len(df),ndigits=2)), "%R:"+str(round(num_r*100/total_testedatb,ndigits=2)), "%S:"+str(round(num_s*100/total_testedatb,ndigits=2)))
                #list of antibiotics passing RIS variation
                min_tested_r,max_tested_r = retrieve_minMaxTestedATB(dict_config_user=configuration_profile, dict_config_default=ACC.dict_configuration_profile_default,
                                                                        param_min=ACC.CONST_VALUE_MIN_RRATE, param_max=ACC.CONST_VALUE_MAX_RRATE)
                min_tested_s,max_tested_s = retrieve_minMaxTestedATB(dict_config_user=configuration_profile, dict_config_default=ACC.dict_configuration_profile_default,
                                                                    param_min=ACC.CONST_VALUE_MIN_SRATE, param_max=ACC.CONST_VALUE_MAX_SRATE)
                lst_col_ris_passC2.append(select_atb_byminMaxTestedATB(str_atb=atb,numerator=num_r,denominator=total_testedatb,min_tested=min_tested_r,max_tested=max_tested_r))
                lst_col_ris_passC2.append(select_atb_byminMaxTestedATB(str_atb=atb,numerator=num_s,denominator=total_testedatb,min_tested=min_tested_s,max_tested=max_tested_s))
        except Exception as e:
            AL.printlog("Error, ANNEX C selecting Antibiotics for profiling: " +  str(e),True,logger)
    lst_col_ris_include_unique = [atb for atb in set(lst_col_ris_include) if atb != ""]
    lst_col_ris_passC1_unique  = [atb for atb in set(lst_col_ris_passC1) if atb != ""] #v3.1 3102
    lst_col_ris_passC2_unique  = [atb for atb in set(lst_col_ris_passC2) if atb != ""] #v3.1 3102
    return lst_col_ris_include_unique, lst_col_ris_passC1_unique, lst_col_ris_passC2_unique #v3.1 3102

#Creating "R---I-S---R..."
#lst_col_ris is list of full antibiotics recommended use from CLSI for that pathogen
#lst_col_ristemp is list of antibiotics for profiling
def create_risprofile(logger,df=pd.DataFrame(), lst_col_ris=[], lst_col_ristemp=[], col_profile="", col_profiletemp=""):
    df[col_profile] = ""
    df[col_profiletemp] = ""
    for atb in lst_col_ris:
        if atb in lst_col_ristemp:
            try:
                df[col_profile] = df[col_profile] + ";" + df[atb].astype(str)
                df[col_profiletemp] = df[col_profiletemp] + ";" + df[atb].astype(str)
            except Exception as e:
                AL.printlog("Error, ANNEX C creating profile column: " + str(e),True,logger)
        else:
            #assign to pos of antibiotic that DO NOT INCLUDED (NC) to profiling >>> need to rename from "NC" to "-" before returning df
            df[col_profile] = df[col_profile] + ";" + "NC"
    # - >>> DO NOT AVAILABLE (NA) or not tested but included for profiling
    #DO NOT INCLUDED (NC) to profiling >>> -
    df[col_profile] = df[col_profile].str.replace("-","NA").str.replace("NC","-").str[1:] #v3.1 3102
    df[col_profiletemp] = df[col_profiletemp].str.replace("-","NA").str[1:] #v3.1 3102
    return df
#Droping RIS duplicated sequences
def drop_dupProfile(df=pd.DataFrame(), col_profile=""):
    return df[col_profile].drop_duplicates(keep="first")
def create_dictforMapProfileID(lst_profile=[], prefixID=""):
    d_profile = {}
    for i in range(len(lst_profile)):
        d_profile[lst_profile[i]] = prefixID+str(i+1) #profile should start from 1
    return d_profile

#####Small functions in the step1.2#####
#Parsing df from the step before grouping isolates by specimen date to satscan_input.csv, satscan_location.csv, and satscan_param.prm.
#Grouping isolates by specimen collection date, prepare location and prm files
#input : excel file of dataframe of deduplicated hospital-origin isolates with profiles
#output: csv files for satscan_input.csv, satscan_location.csv, and satscan_param.prm
def prepare_satscanInput(logger,filename_isolate="", filename_ward="", filename_case="", filename_oriparam="", filename_newparam="", path_input="", path_output="",
                         sh_org="", sh_spc="", start_date="",  end_date="", configuration_prm={}, configuration_default=ACC.dict_configuration_prm_default, str_ward=AC.CONST_VARNAME_WARD, 
                         col_amass=ACC.CONST_COL_AMASSNAME,    col_user=ACC.CONST_COL_USERNAME,        col_resist=ACC.CONST_COL_RESISTPROFILE,
                         col_mrsa=AC.CONST_NEWVARNAME_ASTMRSA_RIS,col_3gc=AC.CONST_NEWVARNAME_AST3GC_RIS,    col_crab=AC.CONST_NEWVARNAME_ASTCBPN_RIS,  col_van=ACC.CONST_NEWVARNAME_ASTVAN_RIS,
                         col_org=AC.CONST_NEWVARNAME_ORG3,    col_oriorg=AC.CONST_VARNAME_ORG,       col_profile=ACC.CONST_COL_PROFILEID, 
                         col_wardname=AC.CONST_VARNAME_WARD ,  col_wardid=AC.CONST_NEWVARNAME_WARDCODE,col_cleandate=AC.CONST_NEWVARNAME_CLEANSPECDATE, 
                         col_testgroup=ACC.CONST_COL_TESTGROUP,col_numcase=ACC.CONST_COL_CASE):
    df = pd.DataFrame()
    directory_path = path_input
    file_name = filename_isolate+"_"+sh_org.upper()+"_"+ACC.dict_spc[sh_spc]+".xlsx"
    file_path = os.path.join(directory_path, file_name)

    if os.path.exists(file_path):
        #For satscan_input.csv
        try:
            df = pd.read_excel(path_input+filename_isolate+"_"+sh_org.upper()+"_"+ACC.dict_spc[sh_spc]+".xlsx").fillna("")
            df = df.loc[:,[col_wardname,col_wardid,col_oriorg,col_org,col_cleandate,col_profile,col_mrsa,col_3gc,col_crab,col_van]]
        except Exception as e:
            AL.printlog("Error, ANNEX C reading deduplicated microbiology file: " +  str(e),True,logger)
        df_ns = pd.DataFrame(columns=[col_testgroup,col_cleandate,col_numcase])
        try:
            d_org = ACC.dict_org[sh_org][0] #get full org name i.e. 'organism_acinetobacter_baumannii'
            for value in ACC.dict_ast[d_org]:
                if sh_org in value:
                    df_temp = df.copy()
                    df_temp[col_resist] = value[-1] #put short resistant name i.e. crab
                    ##Grouping records by cluster_code and specimen_collection_date
                    df_temp = create_clusterCode(df=df_temp, col_testgroup=col_testgroup, col_ward=col_wardid, col_org=col_resist, col_profile=col_profile)
                    df_temp[col_numcase] = 1
                    df_temp_code = df_temp.groupby(by=[col_testgroup,col_cleandate]).sum().reset_index()
                    df_ns = pd.concat([df_ns,df_temp_code],axis=0)
                    df_ns = df_ns.reset_index().loc[:,[col_testgroup,col_numcase,col_cleandate]]
                    df_ns = df_ns.sort_values(by=[col_cleandate]).reset_index().drop(columns=["index"])
                    try:
                        df_ns.to_csv(path_output+filename_case+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",index=False)
                    except Exception as e:
                        AL.printlog("Error, ANNEX C SaTScan input exportation: " +  str(e),True,logger)
                else:
                    pass
        except Exception as e:
            AL.printlog("Error, ANNEX C Case file preparation: " +  str(e),True,logger)
        #For satscan_location.csv
        try:
            prepare_locfile(lst_clustercode=df_ns[col_testgroup].unique(), df_ward_ori=df.loc[:,[col_wardid,col_wardname]], col_wardid=col_wardid, col_testgroup=col_testgroup,  col_wardname=col_wardname, col_profileid=col_profile, col_resistprofile=col_resist, sh_org=sh_org, sh_spc=sh_spc, path_output=path_output)
        except Exception as e:
            AL.printlog("Error, ANNEX C Location file preparation: " +  str(e),True,logger)
        #For satscan_param.prm
        try:
            prepare_prmfile(configuration_prm=configuration_prm, configuration_default=configuration_default,ori_prmfile=filename_oriparam, new_prmfile=filename_newparam, 
                            start_date=start_date, end_date=end_date, path_output=path_output,sh_org=sh_org, sh_spc=sh_spc)
        except Exception as e:
            AL.printlog("Error, ANNEX C Parameter file preparation: " +  str(e),True,logger)
    else:
        pass

    
def read_dictionary(prefix_filename=""):
    df = pd.DataFrame()
    try:
        df = pd.read_excel(prefix_filename+".xlsx")
    except:
        try:
            df = pd.read_csv(prefix_filename+".csv")
        except:
            df = pd.read_csv(prefix_filename+".csv", encoding="windows-1252")
    return df.iloc[:,:2].fillna("")
#Retrieving user value from dictionary_for_microbiology_data.xlsx
def retrieve_uservalue(df, amass_name, col_amass="", col_user=""):
    return df.loc[df[col_amass]==amass_name,:].reset_index().loc[0,col_user]
#Retrieve date from Report1.csv
def retrieve_startEndDate(filename="",col_datafile=ACC.CONST_COL_DATAFILE,col_param=ACC.CONST_COL_PARAM,val_datafile=ACC.CONST_VALUE_DATAFILE,val_sdate=ACC.CONST_VALUE_SDATE,val_edate=ACC.CONST_VALUE_EDATE,col_date=ACC.CONST_COL_DATE):
    df = pd.read_csv(filename)
    start_date = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_sdate),col_date].tolist()[0])
    end_date   = get_dateforprm(df.loc[(df[col_datafile]==val_datafile)&(df[col_param]==val_edate),col_date].tolist()[0])
    return start_date, end_date
#Preparing date from 01 Dec 2022 to 2022/12/01
def get_dateforprm(date=""):
    fmt_date = ""
    month = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
             "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
    for keys,values in month.items():
        if keys in date:
            map_date = date.replace(keys,values).split(" ")
            fmt_date = str(map_date[2])+"/"+str(map_date[1])+"/"+str(map_date[0])
            break
    return fmt_date
#Finding a dictionary key from value
def get_keysfromvalues(str_finder="", d_={}):
    for keys,values in d_.items():
        if values == str_finder:
            return keys
#create a cluster code column
def create_clusterCode(df=pd.DataFrame(), col_testgroup="", col_ward="", col_org="", col_profile=""):
    df[col_testgroup] = df[col_ward] + ";" + df[col_profile] + ";" + df[col_org]
    return df
#Prepare parameter file
def prepare_prmfile(configuration_prm={}, configuration_default={} ,ori_prmfile="",start_date="", end_date="", new_prmfile="", sh_org="", sh_spc="", path_output=""):
    str_wholefile = ""
    file = open(ori_prmfile,"r", encoding="ascii")
    f = file.readline()
    while f != "":  
        for keys,values in configuration_prm.items():
            if keys in f:
                if ("CaseFile=" == keys) or ("CoordinatesFile=" == keys) or ("ResultsFile=" == keys):
                    f = keys+ AC.CONST_PATH_ROOT+values+sh_org+"_"+sh_spc+".csv" + "\n"
                    break
                else:
                    val_ = select_configvalue(configuration_user=configuration_prm,configuration_default=configuration_default,b_satscan=True,param=keys)
                    f = keys + val_ + "\n"
                    break
            else:
                if ("StartDate=" in f) and ("ProspectiveStartDate=" not in f):
                    f = "StartDate="+ start_date + "\n"
                    break
                elif ("EndDate=" in f):
                    f = "EndDate="+ end_date + "\n"
                    break
                else:
                    pass
        str_wholefile = str_wholefile + f
        f = file.readline()
    file.close()
    filename_prm = open(path_output+new_prmfile+"_"+sh_org+"_"+sh_spc+".prm", "w")
    filename_prm.write("".join(str_wholefile))
    filename_prm.close()
#Prepare location file
def prepare_locfile(lst_clustercode=[], df_ward_ori=pd.DataFrame(),
                    col_testgroup="", col_wardid="", col_wardname="", col_profileid="", col_resistprofile="", 
                    sh_org="", sh_spc="", path_output=""):
    df = pd.DataFrame(lst_clustercode,columns=[col_testgroup]).reset_index().drop(columns=["index"])
    df[["x_axis","y_axis"]]=0
    for idx in df.index:
        df.at[idx,"x_axis"]=idx*3
        df.at[idx,"y_axis"]=idx*3
    df[col_testgroup+"_imd"]=df[col_testgroup].str.split(";")
    df[col_wardid]            =df[col_testgroup+"_imd"].str[0]
    df[col_profileid]         =df[col_testgroup+"_imd"].str[1]
    df[col_resistprofile]     =df[col_testgroup+"_imd"].str[2]
    df = df.drop(columns=[col_testgroup+"_imd"])
    df_ward_ori = df_ward_ori.drop_duplicates(keep="first")
    df_2 = df_ward_ori.merge(df, left_on=col_wardid, right_on=col_wardid,how="inner")
    df_2.loc[:,[col_testgroup,"x_axis","y_axis"]].to_csv(path_output+ACC.CONST_FILENAME_LOCATION+"_"+sh_org+"_"+sh_spc+".csv",encoding="ascii",errors="replace",index=False)

#######################################################################
#######################################################################
################# Step2. Performing SaTScan program ###################
#######################################################################
#######################################################################
import os
import subprocess
def call_SaTScan(logger,path_output="",prmfile=""):
    for sh_org in ACC.dict_org.keys():
        for sh_spc in ACC.dict_spc.keys():
            if os.path.exists(path_output+ACC.CONST_FILENAME_INPUT+"_"+sh_org+"_"+sh_spc+".csv"):
                try:
                    subprocess.run(".\Programs\AMASS_amr\SaTScanBatch64.exe -f "+ path_output+prmfile+"_"+sh_org+"_"+sh_spc+".prm", shell=True, text=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
                except subprocess.CalledProcessError as e:
                    error_message = e.stderr
                    AL.printlog("Error, ANNEX C Cluster signals for "+ str(sh_org).upper() + " in " + str(ACC.dict_spc[sh_spc])+": " + str(e.stderr),True,logger)
                else:
                    AL.printlog("Performing cluster signal detection for "+ str(sh_org).upper() + " in " + str(ACC.dict_spc[sh_spc]),False,logger)
            else:
                pass

#######################################################################
#######################################################################
################## Step3. Visualization for AnnexC ####################
#######################################################################
#######################################################################
def prepare_annexc_results(logger,b_wardhighpat=True, num_wardhighpat=2):
    for sh_spc in ACC.dict_spc.keys():
        df_num = pd.DataFrame(0,index=ACC.dict_org.keys(),columns=[ACC.CONST_COL_CASE,ACC.CONST_COL_WARDID])
        for sh_org in ACC.dict_org.keys():
            #read satscan result
            df_re = pd.DataFrame()
            try:
                df_re = read_texttocol(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_RESULT+"_"+sh_org+"_"+sh_spc+".col.txt")
                #ward_015;Profile_7;crab;ward_model
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = df_re.loc[:,ACC.CONST_COL_LOCID].str.split(";", expand=True)
            except Exception as e:
                df_re[[ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_RESISTPROFILE]] = ""

            #formatting datetime
            if len(df_re) > 0:
                df_re_date = AL.fn_clean_date(df=df_re,     oldfield=ACC.CONST_COL_SDATE,cleanfield=ACC.CONST_COL_CLEANSDATE,dformat="",logger="")
                df_re_date = AL.fn_clean_date(df=df_re_date,oldfield=ACC.CONST_COL_EDATE,cleanfield=ACC.CONST_COL_CLEANEDATE,dformat="",logger="")
                df_re_date[ACC.CONST_COL_CLEANSDATE] = df_re_date[ACC.CONST_COL_CLEANSDATE].astype(str)
                df_re_date[ACC.CONST_COL_CLEANEDATE] = df_re_date[ACC.CONST_COL_CLEANEDATE].astype(str)
                df_re_date_pval = format_lancent_pvalue(df=df_re_date,
                                                        col_pval=ACC.CONST_COL_PVAL,
                                                        col_clean_pval=ACC.CONST_COL_CLEANPVAL)
            else:
                df_re_date = df_re.copy()
                df_re_date[[ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE]] = ""

            #preparing date and export results
            select_col = [AC.CONST_VARNAME_ORG,ACC.CONST_COL_WARDID,ACC.CONST_COL_PROFILEID,ACC.CONST_COL_CLEANSDATE,ACC.CONST_COL_CLEANEDATE,
                          ACC.CONST_COL_OBS,ACC.CONST_COL_CLEANPVAL,ACC.CONST_COL_CLEANPVAL+"_lancent"]
            df_re_date_1 = pd.DataFrame(columns=select_col)
            if len(df_re_date) > 0:
                try:
                    df_re_date_1 = df_re_date.copy()
                    df_re_date_1[AC.CONST_VARNAME_ORG]   = df_re_date_1[ACC.CONST_COL_RESISTPROFILE].map(ACC.dict_org)
                    df_re_date_1[[AC.CONST_VARNAME_ORG]] = df_re_date_1[AC.CONST_VARNAME_ORG][0][1]
                    df_re_date_1 = df_re_date_1.loc[:,select_col]
                    df_re_date_2 = df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)<0.05,:]
                    reorder_clusters(df=pd.concat([df_re_date_2,df_re_date_1.loc[df_re_date_1[ACC.CONST_COL_CLEANPVAL].astype(float)>=0.05,:]])).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                                                                                ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                                                                                ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                                                                                ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                                                                                ACC.CONST_COL_CLEANPVAL:ACC.CONST_COL_NEWPVAL,
                                                                                                                                                                ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL+"lancent"}).to_excel(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_ACLUSTER+"_"+str(sh_org).upper()+"_"+str(sh_spc)+".xlsx",index=False,header=True)
                    retrieve_databypval(df=reorder_clusters(df=df_re_date_1),col_pval=ACC.CONST_COL_CLEANPVAL).rename(columns={ACC.CONST_COL_WARDID:ACC.CONST_COL_WARDID,
                                                                                                                                ACC.CONST_COL_CLEANSDATE:ACC.CONST_COL_NEWSDATE,
                                                                                                                                ACC.CONST_COL_CLEANEDATE:ACC.CONST_COL_NEWEDATE,
                                                                                                                                ACC.CONST_COL_OBS:ACC.CONST_COL_NEWOBS, 
                                                                                                                                ACC.CONST_COL_CLEANPVAL+"_lancent":ACC.CONST_COL_NEWPVAL}).drop(columns=[ACC.CONST_COL_CLEANPVAL]).to_excel(AC.CONST_PATH_TEMPWITH_PID+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org).upper()+"_"+str(sh_spc)+".xlsx",index=False,header=True)
                except Exception as e:
                    df_re_date_1.drop(columns=[ACC.CONST_COL_CLEANPVAL+"_temp"])
                    AL.printlog("Error, ANNEX C "+ACC.CONST_FILENAME_PCLUSTER+"_"+str(sh_org)+"_"+str(sh_spc)+".xlsx: " + str(e),True,logger)
            else:
                pass

#####Small functions in the step3#####
def create_df_forstartweekday(s_studydate="2021/01/01", e_studydate="2021/12/31", fmt_studydate="%Y/%m/%d",
                              col_sweekday="startweekday"):
    from datetime import date, timedelta, datetime
    s_studyweek = int(datetime.strptime(s_studydate,fmt_studydate).strftime("%W"))
    s_studyyear = datetime.strptime(s_studydate,fmt_studydate).year
    e_studyyear = datetime.strptime(e_studydate,fmt_studydate).year
    #if study period is not in the first week of that year >>> set the first week
    if s_studyweek != 0:
        s_studyweek = int(datetime.strptime(str(s_studydate),fmt_studydate).strftime("%W"))
    else:
        pass
    #whether study period is over one year >>> include every years during study period
    if e_studyyear > s_studyyear:
        i = int(s_studyyear)
        e_studyweek = int(datetime.strptime(str(e_studydate),fmt_studydate).strftime("%W"))
        while i < e_studyyear:
            e_studyweek = e_studyweek+int(datetime.strptime(str(i)+"/12/31",fmt_studydate).strftime("%W"))
            i += 1
            print ("year: " + str(i))
    else:
        e_studyweek = int(datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
    total_studyday = (datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate) - datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate)).days+1

    lst_startweekday = get_startweekday_forXyear(s_studydate=str(s_studyyear)+"/01/01",
                                                 fmt_studydate  =fmt_studydate,
                                                 total_studyday =total_studyday,
                                                 total_studyweek=e_studyweek-s_studyweek+1)
    df_graph = pd.DataFrame("",
                        index=range(0,int(e_studyweek+1)), 
                        columns=[col_sweekday])
    for idx in df_graph.index:
        df_graph.at[idx,col_sweekday] = str(idx) +" (" + str(lst_startweekday[idx]) + ")"
    return df_graph
def ori_create_df_forstartweekday(s_studydate="2021/01/01", e_studydate="2021/12/31", fmt_studydate="%Y/%m/%d",
                              col_sweekday="startweekday"):
    from datetime import date, timedelta, datetime
    s_studyweek = int(datetime.strptime(s_studydate,fmt_studydate).strftime("%W"))
    s_studyyear = datetime.strptime(s_studydate,fmt_studydate).year
    e_studyyear = datetime.strptime(e_studydate,fmt_studydate).year
    #if study period is not in the first week of that year >>> set the first week
    if s_studyweek != 0:
        s_studyweek = int(datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate).strftime("%W"))
    else:
        pass
    #whether study period is over one year >>> include every years during study period
    if e_studyyear > s_studyyear:
        i = int(s_studyyear)
        e_studyweek = int(datetime.strptime(str(s_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
        while i < e_studyyear:
            e_studyweek = e_studyweek+int(datetime.strptime(str(i)+"/12/31",fmt_studydate).strftime("%W"))
            i += 1
            print ("year: " + str(i))
    else:
        e_studyweek = int(datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate).strftime("%W"))
    total_studyday = (datetime.strptime(str(e_studyyear)+"/12/31",fmt_studydate) - datetime.strptime(str(s_studyyear)+"/01/01",fmt_studydate)).days+1

    lst_startweekday = get_startweekday_forXyear(s_studydate=str(s_studyyear)+"/01/01",
                                                 fmt_studydate  =fmt_studydate,
                                                 total_studyday =total_studyday,
                                                 total_studyweek=e_studyweek-s_studyweek+1)
    df_graph = pd.DataFrame("",
                        index=range(0,int(e_studyweek+1)), 
                        columns=[col_sweekday])
    for idx in df_graph.index:
        df_graph.at[idx,col_sweekday] = str(idx) +" (" + str(lst_startweekday[idx]) + ")"
    return df_graph
#Assigning number of patient and number of wards
def assign_valuetodf(logger,df=pd.DataFrame(),idx_assign="",col_assign="",val_toassign=0):
    try:
        df.at[idx_assign,col_assign] = val_toassign
    except Exception as e:
        AL.printlog("Error, ANNEX C assigning number of patients and number of wards to a dataframe: " + str(e),True,logger)
    return df
#Reordering clusters by 1.total no.observed cases by wards and 2.start signal date
#return a dataframe with ordered clusters
def reorder_clusters(df=pd.DataFrame(),
                     col_wardid=ACC.CONST_COL_WARDID,col_obs=ACC.CONST_COL_OBS,col_date=ACC.CONST_COL_CLEANSDATE):
    df_final = pd.DataFrame()
    df_1     = df.loc[:,[col_wardid,col_obs]]
    df_1.loc[:,col_obs] = df_1.loc[:,col_obs].astype(int)
    lst_ward = df_1.groupby(by=[col_wardid]).sum().sort_values(by=[col_obs],ascending=False).index.tolist()
    for ward in lst_ward:
        df_temp  = df.loc[df[col_wardid]==ward,:].sort_values(by=[col_date,col_obs],ascending=[True,True],na_position="last")
        df_final = pd.concat([df_final,df_temp],axis=0)
    return df_final
def read_texttocol(filename = ""):
    df=pd.DataFrame()
    try:
        file = open(filename, 'r')
        line = file.readline()
        lst_text = []
        while line!="":
            split = [i for i in line.replace("\n","").split(" ") if i != ""]
            lst_text.append(split)
            line = file.readline()
        file.close()
        df = pd.DataFrame(lst_text[1:],columns=lst_text[0]).fillna("NA")
    except:
        pass
    return df
def format_lancent_pvalue(df=pd.DataFrame(),col_pval="",col_clean_pval=""):
    df[col_clean_pval] = ""
    for idx in df.index:
        ori_pval = df.loc[idx,col_pval]
        new_pval_lancent = ""
        new_pval_temp = float(0)
        if float(ori_pval)>=0.995:
            new_pval_temp = round(float(ori_pval),ndigits=2)
            new_pval_lancent = ">0.99"
        elif (float(ori_pval)>0.1) and (float(ori_pval)<0.995):
            new_pval_temp = round(float(ori_pval),ndigits=2)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.01) and (float(ori_pval)<=0.1):
            new_pval_temp = round(float(ori_pval),ndigits=3)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.001) and (float(ori_pval)<=0.01):
            new_pval_temp = round(float(ori_pval),ndigits=4)
            new_pval_lancent = str(new_pval_temp)
        elif (float(ori_pval)>0.0001) and (float(ori_pval)<=0.001):
            new_pval_temp = round(float(ori_pval),ndigits=5)
            new_pval_lancent = str(new_pval_temp)
        # elif (float(ori_pval)>0.00001) and (float(ori_pval)<=0.0001):
        elif (float(ori_pval)<=0.0001):
            new_pval_temp = round(float(ori_pval),ndigits=6)
            new_pval_lancent = "<0.0001"
        df.at[idx,col_clean_pval+"_lancent"] = new_pval_lancent
        df.at[idx,col_clean_pval] = new_pval_temp
    return df
#retrieving records by pval
def retrieve_databypval(df=pd.DataFrame(),col_pval="",pval=0.05):
    return df.loc[df[col_pval].astype(float)<pval,:]
#backup function for several years of startweekday
def get_startweekday_forXyear(s_studydate="2021/01/01",fmt_studydate="%Y/%m/%d",total_studyday=0,total_studyweek=0):
    from datetime import date, timedelta, datetime
    lst_startweek = []
    today  = datetime.strptime(s_studydate, fmt_studydate).date()
    numday = 0
    start_week0 = today - timedelta(days=today.weekday())
    lst_startweek.append(str(start_week0))
    while numday <= total_studyday:
        numday += 7
        start_weekX = start_week0 + timedelta(days=numday)
        lst_startweek.append(str(start_weekX))
    return lst_startweek[:total_studyweek]
#Plotting AnnexC clusters
def plot(logger, df_org=pd.DataFrame(), palette=[], filename="",
         xlabel="Number of week (Start of weekday)",ylabel="Number of patients(n)*",
         min_yaxis=0,max_yaxis=5,step=2):
    import matplotlib.pyplot as plt
    import numpy as np
    try:
        plt.figure()
        df_org.plot(kind='bar', 
                stacked =True, 
                figsize =(20, 10), 
                color =palette, 
                fontsize=16,edgecolor="black",linewidth=2)
        plt.legend(prop={'size': 14},ncol=4)
        plt.ylabel(ylabel, fontsize=16)
        plt.xlabel(xlabel, fontsize=16)
        plt.yticks(np.arange(min_yaxis,max_yaxis+1,step=step), fontsize=16)
        plt.savefig(filename, format='png',dpi=180,transparent=True, bbox_inches="tight")
        plt.close()
        plt.clf
    except Exception as e:
        AL.printlog("Error, ANNEX C Plotting ANNEXC graph: " + str(e),True,logger)


#Reading and preparing satscan_input.csv
def read_satscaninput(filename="", col_code="", col_wardid="", col_profileid="", col_resist="", col_spcdate="", col_numweek="", col_wardprof=""):
    df_input = pd.DataFrame()
    try:
        df_input = pd.read_csv(filename).reset_index().drop(columns=["index"])
        df_input[[col_wardid,col_profileid,col_resist]] = df_input.loc[:,col_code].str.split(";", expand=True)
        df_input[col_numweek] = pd.to_datetime(df_input[col_spcdate]).dt.strftime("%W").astype(int)
        df_input[col_wardprof] = df_input[col_wardid]+";"+df_input[col_profileid]
    except:
        df_input[[col_code,ACC.CONST_COL_CASE,col_spcdate,
                  col_wardid,col_profileid,col_resist,
                  col_numweek,col_wardprof]] = ""
    return df_input

def assign_numcase(df_graph=pd.DataFrame(),df_data=pd.DataFrame(),lst_topward=[],col_wardid="",col_numweek="",col_numcase=""):
    df_graph_new = df_graph.copy()
    df_graph_new[lst_topward] = 0
    for idx in df_data.index:
        if df_data.loc[idx,col_wardid] in lst_topward:
            df_graph_new.at[df_data.loc[idx,col_numweek],df_data.loc[idx,col_wardid]]=df_data.loc[idx,col_numcase]
        else:
            pass
    return df_graph_new

def assign_numcaseforOthers(df_graph=pd.DataFrame(), df_others=pd.DataFrame(), lst_topward=[], col_index="", col_numweek="", col_numcase=""):
    df_graph["Other wards"] = 0
    for idx in df_others.index:
        df_graph.at[df_others.loc[idx,col_numweek],"Other wards"]=df_others.loc[idx,col_numcase]
    df_graph = df_graph.set_index(col_index)
    return df_graph.loc[:,lst_topward+["Other wards"]]

#Selecting dataframe by profile and ddding values of the passed ward/profile to which df for plotting
def select_dfbylist(df=pd.DataFrame(), lst_lookingfor=[], col_lookingfor="", lst_groupby=[]):
    df_select = pd.DataFrame()
    if len(lst_lookingfor) > 0:
        df_select = df.loc[df[col_lookingfor].isin(lst_lookingfor)].groupby(by=lst_groupby).sum().reset_index()
    else:
        df_select = df
    return df_select
