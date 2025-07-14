#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** MAIN ANALYSIS (SECTION 1-6, ANEX A                                                              ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate AMR surveillance reports, Supplementary data indicators reports, and Data verification logfile reports systematically.
# Code is rewrite from R code developed by Cherry Lim in AMASS version 2.0.15
# @author: PRAPASS WANNAPINIJ
# Created on: 20 Oct 2022
# Last update on: 30 APR 2025

import os
#import sys
#import csv
#import logging #for creating error_log
import pandas as pd #for creating and manipulating dataframe
import numpy as np #for creating arrays and calculations
from datetime import datetime
#from AMASS_3_amr_commonlib import * #for importing amr functions
#from xlsx2csv import Xlsx2csv

import calendar as cld
import psutil,gc
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL
import AMASS_annex_b_analysis as ANNEX_B
import AMASS_amr_report_new as AMR_REPORT_NEW
import AMASS_supplementary_report as SUP_REPORT
import AMASS_amr_commonlib_report as ARC
import AMASS_amr_commonlib_genv2compatible as V2


from scipy.stats import norm # -> must moveto common lib

bisdebug = True
# Function for get lower uper CI -> must moveto common lib
# to calculate lower 95% CI
def fn_wilson_lowerCI(x, n, conflevel, decimalplace):
    zalpha = abs(norm.ppf((1-conflevel)/2))
    phat = x/n
    bound = (zalpha*((phat*(1-phat)+(zalpha**2)/(4*n))/n)**(1/2))/(1+(zalpha**2)/n)
    midpnt = (phat+(zalpha**2)/(2*n))/(1+(zalpha**2)/n)
    lowlim = round((midpnt - bound)*100, decimalplace)
    return lowlim

# to calculate upper 95% CI
def fn_wilson_upperCI(x, n, conflevel, decimalplace):
    zalpha = abs(norm.ppf((1-conflevel)/2))
    phat = x/n
    bound = (zalpha*((phat*(1-phat)+(zalpha**2)/(4*n))/n)**(1/2))/(1+(zalpha**2)/n)
    midpnt = (phat+(zalpha**2)/(2*n))/(1+(zalpha**2)/n)
    uplim = round((midpnt + bound)*100, decimalplace)
    return uplim
def correct_digit(val):
    nval = val
    try:
        if float(nval) < 0.05:
            nval = 0
        elif float(nval) >= 0.95:
            nval= int(round(nval))
        else:
            pass 
    except:
        pass
    return str(nval)
# function that print obj if in debug mode (bisdebug = true)
def printdebug(obj) :
    try:
        if bisdebug :
            print(obj)
    except Exception as e: 
        print(e)
def check_config(df_config, str_process_name):
    #Checking process is either able for running or not
    #df_config: Dataframe of config file
    #str_process_name: process name in string fmt
    #return value: Boolean; True when parameter is set "yes", False when parameter is set "no"
    config_lst = df_config.iloc[:,0].tolist()
    result = ""
    if df_config.loc[config_lst.index(str_process_name),"Setting parameters"] == "yes":
        result = True
    else:
        result = False
    return result


def clean_hn(df,oldfield,newfield,logger) :
    df[newfield] = df[oldfield].fillna("").astype("string").str.strip()
    try:
        df[newfield] = df[newfield].str.replace(" ","").str.lower()
        #Case read and add .0 at the end of HN as number 
        df[newfield] = df[newfield].str.replace(r'\.0$','',regex=True)
    except Exception as e: # work on python 3.x
        #print(e)
        AL.printlog("Warning : unable to clean ',' and hn '.0' bug convertion", False, logger)
    return df
def caldays(df,coldate,dbaselinedate) :
    return (df[coldate] - dbaselinedate).dt.days
def fn_allmonthname():
    month = 1
    return [cld.month_name[(month % 12 + i) or month] for i in range(12)]
def get_randomstring_data(temp_df,sfield,iRow):
    try:
        sstr = ""
        if len(temp_df) > 0:
            temp_df = temp_df.loc[np.random.permutation(temp_df.index)[:iRow if len(temp_df) > iRow else len(temp_df)]]
            for index, row in temp_df.iterrows():
                sstr = sstr + (', ' if len(sstr) > 0 else '') + str(row[sfield])
        return sstr
    except:
        return ""
def check_hn_format(orgdf,hnfield,logger):
    df2 = pd.DataFrame(columns =["str_length","counts","percent"]) 
    try:
        df = orgdf[[hnfield]].fillna('')
        df["str_length"] = df[hnfield].str.len()
        df2 = df.groupby(["str_length"]).size().reset_index(name='counts')
        df2["percent"] = round(df2['counts']/df2['counts'].sum(),4)*100
    except Exception as e:
        AL.printlog("Warning : create data varidation log for HN charecter length distribution", False, logger)
        logger.exception(e)
    return df2
def check_date_format(orgdf,oldfield,logger):
    try:
        isalreadydatecol = pd.api.types.is_datetime64_any_dtype(orgdf[oldfield])
    except:
        isalreadydatecol = False
    df = orgdf[[oldfield]]
    df_logformat = pd.DataFrame(columns =["dateformat","exampledate"]) 
    if isalreadydatecol:
        onew_row = {"dateformat":'Dates are in excel standard date format : ',"exampledate":get_randomstring_data(df,oldfield,3)}   
        df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
    else:
        sfield = oldfield + "amasschk"
        cft_1 = sfield + "_amassd1"
        cft_2 = sfield + "_amassd2"
        try:
            df[sfield] = df[oldfield].astype("string")
            df[sfield] = df[sfield].str.replace('-', '/', regex=False)
            df[cft_1] = df[sfield].str.split("/", n = 2, expand = True)[0]
            df[cft_1] = pd.to_numeric(df[cft_1],downcast='signed',errors='coerce').fillna(0)
            df[cft_2] = df[sfield].str.split("/", n = 2, expand = True)[1]
            df[cft_2] = pd.to_numeric(df[cft_2],downcast='signed',errors='coerce').fillna(0)
            temp_df = df.loc[(df[cft_1]>12) & (df[cft_1]<32)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'DMY',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_2]>12) & (df[cft_2]<32)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'MDY',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_1]>31)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'YMD',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            temp_df = df.loc[(df[cft_1]==0)]
            if len(temp_df) > 0:
                onew_row = {"dateformat":'others defined',"exampledate":get_randomstring_data(temp_df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
        except:
            try:
                onew_row = {"dateformat":'others',"exampledate":get_randomstring_data(df,oldfield,3)}   
                df_logformat = pd.concat([df_logformat,pd.DataFrame([onew_row])],ignore_index = True)
            except:
                pass
    return df_logformat
def clean_date(df,oldfield,cleanfield,dformat,logger):
    return AL.fn_clean_date(df,oldfield,cleanfield,dformat,logger)
def fn_clean_date_andcalday_year_month(df,oldfield,cleanfield,caldayfield,calyearfield,calmonthfield, CDATEFORMAT, ORIGIN_DATE,logger) :
    try:
        bisok = True
        try:
            if oldfield in df.columns:
                df  = clean_date(df, oldfield, cleanfield, CDATEFORMAT,logger)
                if caldayfield!="" : df[caldayfield] = caldays(df, cleanfield, ORIGIN_DATE)
                if calyearfield!="": df[calyearfield] = df[cleanfield].dt.strftime("%Y")
                if calmonthfield!="": df[calmonthfield] = df[cleanfield].dt.strftime("%B")
            else:
                AL.printlog("Warning : No field : " + oldfield, False, logger)
                bisok = False
        except Exception as e:
            AL.printlog("Error : unable to convert and calculate day year month for " + oldfield, False, logger)
            logger.exception(e)
            bisok = False
        if bisok == False:
            df[cleanfield] = np.nan
            if caldayfield!="": df[caldayfield] = np.nan
            if calyearfield!="":df[calyearfield] = np.nan
            if calmonthfield!="": df[calmonthfield] = np.nan
        return df  
    except:
        return df      
def fn_notindata(df_source,df_indi,sfield_source,sfield_indi,bindidropdup=False):
    df1=pd.DataFrame()
    df2=pd.DataFrame()
    try:
        df1[sfield_source] =df_source[sfield_source]
        if bindidropdup:
            df2[sfield_indi+"_indi"] = df_indi[sfield_indi].drop_duplicates()
        else:
            df2[sfield_indi+"_indi"] = df_indi[sfield_indi]
        df1 = df1.merge(df2, how="left", left_on=sfield_source, right_on=sfield_indi+"_indi",suffixes=("", "_indi"))
        df1[sfield_indi+"_indi"] = df1[sfield_indi+"_indi"] .fillna("")
        df1=df1[df1[sfield_indi+"_indi"]==""]
    except:
        pass
    return df1
def fn_mergededup_hospmicro(df_micro,df_hosp,bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger,df_list_matchrid,list_specdate_tolerance) :
    df_merged = pd.DataFrame()
    bMergedsuccess = False
    sErrorat = ""
    # Merge if got same column name in both (such as date_of_admission) it will use the date_of-admission from micro dataframe
    if bishosp_ava:
        try:           
            sErrorat = "(do merge)"
            df_merged = df_micro.copy(deep=True)
            try:
                micro_hn_dtype = df_merged.dtypes[AC.CONST_NEWVARNAME_HN]
                hosp_hn_dtype = df_hosp.dtypes[AC.CONST_NEWVARNAME_HN_HOSP]
                AL.printlog("HN column data type for micro is " + str(micro_hn_dtype),False,logger)
                AL.printlog("HN column data type for hosp data is " + str(hosp_hn_dtype),False,logger)
            except:
                pass
            if micro_hn_dtype != hosp_hn_dtype :
                AL.printlog("HN column data type different between micro and hosp data",False,logger)

            df_merged = df_merged.merge(df_hosp, how="inner", left_on=AC.CONST_NEWVARNAME_HN, right_on=AC.CONST_NEWVARNAME_HN_HOSP,suffixes=("", AC.CONST_MERGE_DUPCOLSUFFIX_HOSP))
            df_list_matchrid[AC.CONST_NEWVARNAME_MICROREC_ID] =df_merged[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates()

            sErrorat = "(do clean date)"
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_ADMISSIONDATE, AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_DAYTOADMDATE, AC.CONST_NEWVARNAME_ADMYEAR, AC.CONST_NEWVARNAME_ADMMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_DISCHARGEDATE, AC.CONST_NEWVARNAME_CLEANDISDATE, AC.CONST_NEWVARNAME_DAYTODISDATE, AC.CONST_NEWVARNAME_DISYEAR, AC.CONST_NEWVARNAME_DISMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_merged = fn_clean_date_andcalday_year_month(df_merged, AC.CONST_VARNAME_BIRTHDAY, AC.CONST_NEWVARNAME_CLEANBIRTHDATE, AC.CONST_NEWVARNAME_DAYTOBIRTHDATE, "","", AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            #cal end date to get admission period
            df_merged[AC.CONST_NEWVARNAME_DAYTOENDDATE] = df_merged[AC.CONST_NEWVARNAME_DAYTODISDATE]
            df_merged[AC.CONST_NEWVARNAME_DAYTOSTARTDATE] = df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]
            try:
                calcolmaxdayinclude(df_merged,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_merged,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOENDDATE,AC.CONST_ORIGIN_DATE,"",logger)   
                calcolmindayinclude(df_merged,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_merged,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOSTARTDATE,AC.CONST_ORIGIN_DATE,"",logger)  
            except Exception as e:
                AL.printlog("Warning : calculate min max include date (Merged data)"  + str(e),True,logger)
                logger.exception(e)
                pass
            # Remove unmatched
            sErrorat = "(do remove unmatch spec date vs admission period)"
            iallowspecdate_before = 0
            iallowspecdate_after = 0
            try:
                iallowspecdate_before = int(list_specdate_tolerance[0]) # value in configuration is negative value such as -1 = 1 day before
                iallowspecdate_after = int(list_specdate_tolerance[1])
            except Exception as e:
                AL.printlog("Warning : get number of days allow specimen date to be before and/or after adminission period"  + str(e),True,logger)
                logger.exception(e)
            df_merged = df_merged[(df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]>=(df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]+iallowspecdate_before)) & (df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE]<=(df_merged[AC.CONST_NEWVARNAME_DAYTOENDDATE]+iallowspecdate_after))]
            # Translate gender, age year, age cat data
            sErrorat = "(do add columns)"
            # Translate gender, age year, age cat data
            df_merged[AC.CONST_NEWVARNAME_GENDERCAT] = np.nan
            try:
                df_merged[AC.CONST_NEWVARNAME_GENDERCAT] = df_merged[AC.CONST_VARNAME_GENDER].astype("string").str.strip().map(dict_gender_datavaltoamass)
            except Exception as e:
                AL.printlog("Warning : Unable to identify gender data : "  + str(e),True,logger)
                logger.exception(e)
            # CO/HO (If want to support case micro data have admission date move it outside of bishosp_ava condition)
            if AC.CONST_VARNAME_COHO in df_merged.columns:
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = df_merged[AC.CONST_VARNAME_COHO].astype("string").str.strip().map(dict_inforg_datavaltoamass)
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.select([(df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] == AC.CONST_DICT_COHO_CO),(df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] == AC.CONST_DICT_COHO_HO)],
                                                                            [0,1],
                                                                            default=np.nan)
            else:
                df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL] = np.select([((df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE] - df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) < 2),
                                                                         ((df_merged[AC.CONST_NEWVARNAME_DAYTOSPECDATE] - df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE]) >= 2)],
                                                                        [0,1],
                                                                        default=np.nan)     
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS]==AC.CONST_DICT_COHO_AVAIL][AC.CONST_DICTCOL_DATAVAL].values[0] == "yes" :
               df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] =  df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS]
            else:
               df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] =  df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL]  
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] =np.nan
            try:
                df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] = df_merged[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").str.strip().map(dict_died_datavaltoamass).fillna(AC.CONST_ALIVE_VALUE) # From R code line 1154
            except Exception as e:
                AL.printlog("Warning : Unable to identify outcome (died,alive) : "  + str(e),True,logger)
                logger.exception(e)
            sErrorat = "(do change columns to category type)"
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] = df_merged[AC.CONST_NEWVARNAME_DISOUTCOME].astype("category")
            #Managing ward that may be from either micro or hosp
            try:
                df_merged[AC.CONST_NEWVARNAME_WARDCODE].fillna(df_merged[AC.CONST_NEWVARNAME_WARDCODE_HOSP], inplace=True)
                df_merged[AC.CONST_NEWVARNAME_WARDTYPE].fillna(df_merged[AC.CONST_NEWVARNAME_WARDTYPE_HOSP], inplace=True)
                df_merged.drop(AC.CONST_NEWVARNAME_WARDCODE_HOSP, axis=1, inplace=True) 
                df_merged.drop(AC.CONST_NEWVARNAME_WARDTYPE_HOSP, axis=1, inplace=True)  
                df_merged[AC.CONST_NEWVARNAME_WARDCODE].fillna(AC.CONST_WARD_ID_WARDNOTINDICT, inplace=True)
                df_merged[AC.CONST_NEWVARNAME_WARDTYPE].fillna("", inplace=True)
            except Exception as e: # work on python 3.x
                AL.printlog("Warning unable to manage ward info from micro and hospital admission data : " + sErrorat + " : " + str(e),False,logger)
                logger.exception(e)
            #Build 3027 move ageyear calculate after clean
            df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = 0
            try:
                if df_dict[df_dict[AC.CONST_DICTCOL_AMASS]=='age_year_available'][AC.CONST_DICTCOL_DATAVAL].values[0] == "yes" :
                #if dict_amasstodataval['age_year_available'].lower() == "yes":
                    df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_VARNAME_AGEY].apply(pd.to_numeric,errors='coerce')
                    AL.printlog("Warning : age from age year column ",False,logger)
                elif df_dict[df_dict[AC.CONST_DICTCOL_AMASS]=='birthday_available'][AC.CONST_DICTCOL_DATAVAL].values[0] == "yes" :
                #elif dict_amasstodataval['birthday_available'].lower() == "yes":
                    df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] =  (df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE])/365.25
                    AL.printlog("Warning : age from calculate birthday and admission date ",False,logger)
                df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].apply(np.floor,errors='coerce')
            except:
                AL.printlog("Warning : unable to calculate age year ",False,logger)
            bMergedsuccess = True
        except Exception as e: # work on python 3.x
            AL.printlog("Failed merge hosp and micro data on " + sErrorat + " : " + str(e),True,logger)
            logger.exception(e)
            pass
    if (bishosp_ava==False) or (bMergedsuccess==False) :
        try:
            AL.printlog("No hosp data/merged error, merged data is micro data",False,logger)
            df_merged = df_micro.copy(deep=True)
            df_merged[AC.CONST_NEWVARNAME_HN_HOSP] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANADMDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTOADMDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_ADMYEAR] = np.nan
            df_merged[AC.CONST_NEWVARNAME_ADMMONTHNAME] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANDISDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTODISDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISYEAR] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISMONTHNAME] = np.nan
            df_merged[AC.CONST_NEWVARNAME_CLEANBIRTHDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE] = np.nan
            df_merged[AC.CONST_NEWVARNAME_GENDERCAT] = np.nan
            df_merged[AC.CONST_NEWVARNAME_AGEYEAR] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMHOS] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FROMCAL] = np.nan
            df_merged[AC.CONST_NEWVARNAME_COHO_FINAL] = np.nan
            df_merged[AC.CONST_NEWVARNAME_DISOUTCOME] =np.nan
        except Exception as e: # work on python 3.x
            AL.printlog("Failed case no hosp data/merged error, merged data is micro data : " + str(e),True,logger)
            logger.exception(e)
    return df_merged
def calcolmindayinclude(dfm,scolm,dfh,scolh,scolstartdate,dorgdate,ddefault,logger) :
    try:
        dmin_data_include = dfm[scolm].min()
        dtemp = dfh[scolh].min()
        if dtemp > dmin_data_include:
            dmin_data_include = dtemp
        idaytomin_date_include = (dmin_data_include - dorgdate).days
        AL.printlog("Min date include = " + str(dmin_data_include) + " which is days " + str(idaytomin_date_include),False,logger)
        dfh[scolstartdate] = dfh[scolstartdate].fillna(idaytomin_date_include)
        dfh.loc[dfh[scolstartdate] < idaytomin_date_include, scolstartdate] = idaytomin_date_include
        #reset those adm date is null back to null value
        dfh.loc[dfh[scolh].isnull(), scolstartdate] = np.nan
        return dmin_data_include
    except Exception as e:
        AL.printlog("Warning : Fail to calculate/set value start date data include: " +  str(e),False,logger)
        logger.exception(e)
        return ddefault
def calcolmaxdayinclude(dfm,scolm,dfh,scolh,scolendate,dorgdate,ddefault,logger) :
    try:
        dmax_data_include = dfm[scolm].max()
        dtemp = dfh[scolh].max()
        if dtemp < dmax_data_include:
            dmax_data_include = dtemp
        idaytomax_date_include = (dmax_data_include - dorgdate).days
        AL.printlog("Max date include = " + str(dmax_data_include) + " which is days " + str(idaytomax_date_include),False,logger)
        dfh[scolendate] = dfh[scolendate].fillna(idaytomax_date_include)
        dfh.loc[dfh[scolendate] > idaytomax_date_include, scolendate] = idaytomax_date_include
        #reset those adm date is null back to null value
        dfh.loc[dfh[scolh].isnull(), scolendate] = np.nan
        return dmax_data_include
    except Exception as e:
        AL.printlog("Warning : Fail to calculate/set value end date data include: " +  str(e),False,logger)
        logger.exception(e)
        return ddefault
# Save temp file if in debug mode
def debug_savecsv(df,fname,bdebug,iquotemode,logger)  :
    if bdebug :
        AL.fn_savecsv(df,fname,iquotemode,logger)
        #return True
# Filter orgcat before dedup (For merge data)
def fn_deduplicatebyorgcat_hospmico(df,colname,orgcat) :
    #return fn_deduplicatedata(df.loc[df[colname]==orgcat],[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AMR,AC.CONST_NEWVARNAME_AMR_TESTED,AC.CONST_NEWVARNAME_CLEANADMDATE],[True,True,False,False,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
    #Change in 3.1
    return fn_deduplicatedata(df.loc[df[colname]==orgcat],
                              [AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AST_R,AC.CONST_NEWVARNAME_AST_I,AC.CONST_NEWVARNAME_AST_TESTED,AC.CONST_NEWVARNAME_CLEANADMDATE],
                              [True,True,False,False,False,True],
                              "last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
# Filter orgcat before dedup
def fn_deduplicatebyorgcat(df,colname,orgcat) :
    #return fn_deduplicatedata(df.loc[df[colname]==orgcat],[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AMR,AC.CONST_NEWVARNAME_AMR_TESTED],[True,True,False,False],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
    #Change in 3.1
    return fn_deduplicatedata(df.loc[df[colname]==orgcat],
                              [AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_AST_R,AC.CONST_NEWVARNAME_AST_I,AC.CONST_NEWVARNAME_AST_TESTED],
                              [True,True,False,False,False],
                              "last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
# deduplicate data by sort the data by multiple column (First order then second then..) and then select either first or last of each group of the first order and/or second and/or..
def fn_deduplicatedata(df,list_sort,list_order,na_posmode,list_dupcolchk,keepmode) :
    return df.sort_values(by = list_sort, ascending = list_order, na_position = na_posmode).drop_duplicates(subset=list_dupcolchk, keep=keepmode)
def fn_getunknownbyorg(sorg,df_COHO,df_all,sidcol,logger):
    df_unknown = pd.DataFrame()
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at list unkown origin for " + sorg + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
        ssuffix_coho = "_COHO"
        sidcol_coho = sidcol + ssuffix_coho
        temp_df = df_COHO.copy(deep=True)
        temp_df[sidcol_coho] = temp_df[sidcol]
        temp_df = temp_df[[sidcol_coho]]
        df_unknown = df_all.copy(deep=True)
        df_unknown = df_unknown.merge(temp_df, how="left", left_on=sidcol, right_on=sidcol_coho,suffixes=("", ssuffix_coho))
        df_unknown[sidcol_coho].fillna("", inplace=True)
        df_unknown = df_unknown[df_unknown[sidcol_coho]==""]
    except Exception as e:
        AL.printlog("Error : Fail to extract unknown origin deduplicated microbiology record for " + sorg + ": " +  str(e),True,logger)
        logger.exception(e)
    return df_unknown
def sub_printprocmem(sstate,logger) :
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at state " +sstate + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
    except:
        AL.printlog("Error get process memory usage at " + sstate,True,logger)
#Ward dictionary function
def fn_clean_ward(df,scol_wardindict,scol_wardid,scol_wardtype,path,f_dict_ward,logger):
    bOK = False
    if AL.checkxlsorcsv(path,f_dict_ward):
        try:
            df_dict_ward = AL.readxlsorcsv_noheader(path,f_dict_ward, [AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"WARDTYPE","REQ","EXPLAINATION"],logger)
            df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip() != ""]
            if len(df_dict_ward[df_dict_ward [AC.CONST_DICTCOL_AMASS] == scol_wardindict]) > 0:
                scol_wardorg = df_dict_ward[df_dict_ward [AC.CONST_DICTCOL_AMASS] == scol_wardindict].iloc[0][AC.CONST_DICTCOL_DATAVAL]
                df_dict_ward = df_dict_ward[df_dict_ward[AC.CONST_DICTCOL_AMASS].str.startswith("ward_")]
                if scol_wardorg in df.columns:
                    tempdict = pd.Series(df_dict_ward[AC.CONST_DICTCOL_AMASS].values,index=df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
                    df[scol_wardid] = df[scol_wardorg].astype("string").str.strip().map(tempdict)
                    tempdict = pd.Series(df_dict_ward["WARDTYPE"].values,index=df_dict_ward[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
                    df[scol_wardtype] = df[scol_wardorg].astype("string").str.strip().map(tempdict)
                    df.rename(columns={scol_wardorg:scol_wardindict}, inplace=True)
                    bOK = True
                else:
                    AL.printlog("Warning : Fail to convert ward data: " + "No ward column as specify in dictionary of ward",False,logger)
            else:
                AL.printlog("Warning : No ward configuration : " + scol_wardindict + " : in dictionary of ward",False,logger)
        except Exception as e:
            AL.printlog("Warning : Fail to convert ward data: " +  str(e),False,logger)
            logger.exception(e) 
    else:
        AL.printlog("Note : No dictionary for ward file",False,logger)  
    if bOK == False:
        df[scol_wardid] = np.nan
        df[scol_wardtype] = np.nan
    try:
        df[scol_wardid].astype("category")
        df[scol_wardtype].astype("category")
    except:
        pass
    return bOK    
def mainloop() :    
    dict_progvar = {}  
    df_micro_ward = pd.DataFrame()
    ## Init log 
    logger = AL.initlogger('AMR anlaysis',"./log_amr_analysis.txt")
    AL.printlog("AMASS version : " + AC.CONST_SOFTWARE_VERSION_FORLOG,False,logger)
    AL.printlog("Pandas library version : " + str(pd.__version__),False,logger)
    AL.printlog("Start AMR analysis: " + str(datetime.now()),False,logger)
    ## Date format
    fmtdate_text = "%d %b %Y"
    try:
        #dict_cmdarg  = AL.getcmdargtodict(logger)
        bIsDoAnnexC = True
        """
        try:
            if dict_cmdarg['annex_c'] == "no":
                bIsDoAnnexC = False
                AL.printlog("Note : Command line specify not do Annex C",False,logger)
        except:
            pass
        """
        try:
            config=AL.readxlsxorcsv(AC.CONST_PATH_ROOT + "Configuration/", "Configuration",logger)
            bIsDoAnnexC = ARC.check_config(config, "amr_surveillance_annexC")
        except:
            pass
        # If folder doesn't exist, then create it.
        if not os.path.isdir(AC.CONST_PATH_RESULT) : os.makedirs(AC.CONST_PATH_RESULT)
        path_repwithPID = AC.CONST_PATH_REPORTWITH_PID
        path_variable = AC.CONST_PATH_VAR
        path_input = AC.CONST_PATH_ROOT
        
        sub_printprocmem("Start main loop",logger)
        bishosp_ava = AL.checkxlsorcsv(path_input,"hospital_admission_data")
        # Import dictionary
        s_dict_column =[AC.CONST_DICTCOL_AMASS,AC.CONST_DICTCOL_DATAVAL,"REQ","EXPLAINATION"]
        df_dict_micro = AL.readxlsorcsv_noheader(path_input,"dictionary_for_microbiology_data",s_dict_column,logger)
        df_dict_micro.loc[df_dict_micro[s_dict_column[1]].isnull() == True, s_dict_column[1]] = AC.CONST_DICVAL_EMPTY
        df_dict_hosp = pd.DataFrame()
        if bishosp_ava:
            df_dict_hosp = AL.readxlsorcsv_noheader(path_input,"dictionary_for_hospital_admission_data",s_dict_column,logger)
            df_dict_hosp.loc[df_dict_hosp[s_dict_column[1]].isnull() == True, s_dict_column[1]] = AC.CONST_DICVAL_EMPTY
        #Combine data dict
        df_dict = pd.DataFrame()
        if bishosp_ava:
            df_combine_dict = [df_dict_micro.iloc[1:,0:2],df_dict_hosp.iloc[1:,0:2]]
            df_dict = pd.concat(df_combine_dict)
        else:
            df_dict = df_dict_micro.iloc[1:,0:2] 
        AL.printlog("Complete load dictionary file",False,logger)  
        ## Get HN variable from dict
        hn_col_micro = ""
        try:
            hn_col_micro = df_dict_micro[df_dict_micro[AC.CONST_DICTCOL_AMASS] == 'hospital_number'].iloc[0][AC.CONST_DICTCOL_DATAVAL]
        except:
            hn_col_micro = ""
        hn_col_hosp = ""
        if bishosp_ava:
            try:
                hn_col_hosp = df_dict_hosp[df_dict_hosp[AC.CONST_DICTCOL_AMASS] == 'hospital_number'].iloc[0][AC.CONST_DICTCOL_DATAVAL]
            except:
                hn_col_hosp =""
        ## Import data 
        df_micro = pd.DataFrame()
        df_hosp = pd.DataFrame()
        df_hosp_formerge = pd.DataFrame()
        df_config = AL.readxlsxorcsv(path_input +"Configuration/", "Configuration",logger)
        if check_config(df_config, "amr_surveillance_function") :
            #import microbiology
            if AL.checkxlsorcsv(path_input,"microbiology_data_reformatted") :
                df_micro = AL.readxlsxorcsv_forcehntostring(path_input,"microbiology_data_reformatted",hn_col_micro,logger)
            else :
                df_micro = AL.readxlsxorcsv_forcehntostring(path_input,"microbiology_data",hn_col_micro,logger)
            df_micro_annexb = df_micro.copy(deep=True)
        AL.printlog("Succesful read micro data file with " + str(len(df_micro)) + " records",False,logger)  
        #Special task for log ast --------------------------------------------------------------------------------------------------
        #"Save the log_ast.xlsx for data validation file"
        temp_df = pd.DataFrame(columns =["Antibiotics","frequency_raw"]) 
        for scolname in df_micro:
            try:   
                n_row = len(df_micro[(df_micro[scolname].isnull() == False) & (df_micro[scolname] != "")])
                onew_row = {"Antibiotics":scolname,"frequency_raw":str(n_row)}     
                temp_df = pd.concat([temp_df,pd.DataFrame([onew_row])], ignore_index = True)
            except Exception as e: 
                print(e)    
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_ast.xlsx", logger):
            AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_ast.xlsx",False,logger)  
        del temp_df
        gc.collect()
        if bishosp_ava:
            df_hosp = AL.readxlsxorcsv_forcehntostring(path_input,"hospital_admission_data",hn_col_hosp,logger)
            AL.printlog("Succesful read hospital data file with " + str(len(df_hosp)) + " records",False,logger)  
        else:
            AL.printlog("Note : No hospital data file",False,logger)  
        bishosp_ava = not df_hosp.empty       
        # The following part may be optional  --------------------------------------------------------------------------------------------------
        # Export the list of variables
        if not os.path.exists(path_variable):
            os.mkdir(path_variable)
        list_micro_val = list(df_micro.columns)
        df_micro_val = pd.DataFrame(list_micro_val)
        df_micro_val.columns = ['variables_micro']
        if not AL.fn_savexlsx(df_micro_val, path_variable + "variables_micro.xlsx", logger):
            AL.printlog("Warning: Cannot save xlsx file : " + path_variable + "variables_micro.xlsx",False,logger)
            try:
                df_micro_val.to_csv(path_variable + "variables_micro.csv",encoding='utf-8',index=False)
            except:
                AL.printlog("Warning: Cannot save csv file : " + path_variable + "variables_micro.csv",False,logger)
        if bishosp_ava:
            df_hosp_val = pd.DataFrame(list(df_hosp.columns))
            df_hosp_val.columns = ['variables_hosp']
            if not AL.fn_savexlsx(df_hosp_val, path_variable + "variables_hosp.xlsx", logger):
                AL.printlog("Warning: Cannot save xlsx file : " + path_variable + "variables_hosp.xlsx",False,logger)
                try:
                    df_hosp_val.to_csv(path_variable + "variables_hosp.csv",encoding='utf-8',index=False)
                except:
                    AL.printlog("Warning: Cannot save csv file : " + path_variable + "variables_hosp.csv",False,logger)
        # --------------------------------------------------------------------------------------------------------------------------------------
        
        #Data validation log for df_dict duplicated value
        try:
            temp_df = df_dict[df_dict[AC.CONST_DICTCOL_DATAVAL].str.strip() != ""]
            temp_df = temp_df.groupby([AC.CONST_DICTCOL_DATAVAL]).size().reset_index(name='counts')
            temp_df = temp_df[temp_df['counts'] > 1]
            temp_df =temp_df.merge(df_dict, how="inner", left_on=AC.CONST_DICTCOL_DATAVAL, right_on=AC.CONST_DICTCOL_DATAVAL,suffixes=("", "_NOTUSED"))
            list_notconsiderdup = ['no','No','NO','yes','Yes','YES','xxx_Can be changed in the dictionary_of_variable_data.csv_xxx',
                                   'Data values of the variable recorded for "organism" in your microbiology data file',
                                   'Data values of the variable recorded in your microbiology data file']
            temp_df = temp_df[~temp_df[AC.CONST_DICTCOL_DATAVAL].isin(list_notconsiderdup)]
            temp_df_h = temp_df[temp_df[AC.CONST_DICTCOL_AMASS] == 'hospital_number']
            if len(temp_df_h) > 0:
                temp_df = temp_df[~temp_df[AC.CONST_COL_AMASS].isin(['hospital_number'])]
            if len(temp_df) > 0:
                temp_df = temp_df[[AC.CONST_DICTCOL_DATAVAL,AC.CONST_DICTCOL_AMASS]]
            else:
                temp_df = pd.DataFrame(columns=[AC.CONST_DICTCOL_DATAVAL,AC.CONST_DICTCOL_AMASS])
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_dup.xlsx", logger):
                AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_dup.xlsx",False,logger)  
            del temp_df
        except Exception as e: 
            AL.printlog("Warning : check for ducplicate value in both dictionary: " +  str(e),False,logger)   
        
        AL.printlog("Complete check and log dictionary file variables",False,logger)  
        #Reverse order of dict -> to be the same as R when map
        df_dict = df_dict[::-1].reset_index(drop = True) 
        bisabom = True
        try:
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "acinetobacter_spp_or_baumannii"].iloc[0][AC.CONST_DICTCOL_DATAVAL] != "organism_acinetobacter_baumannii":
                bisabom = False
        except Exception as e: 
            AL.printlog("Warning : unable to read acinetobacter_spp_or_baumannii configuration from dictionary: " +  str(e),False,logger)   
        bisentspp = False
        try:
            if df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "enterococcus_spp_or_faecalis_and_faecium"].iloc[0][AC.CONST_DICTCOL_DATAVAL] == "organism_enterococcus_spp":
                bisentspp = True
        except Exception as e: 
            AL.printlog("Warning : unable to read enterococcus_spp_or_faecalis_and_faecium configuration from dictionary: " +  str(e),False,logger)  
        # Grep function in R - may be the following Python code do it wrongly --- this may need to remove by changing the data dictionary
        if bisabom == False:
            df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_acinetobacter"),AC.CONST_DICTCOL_AMASS] = "organism_acinetobacter_spp"
        if bisentspp == True:
            df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_enterococcus"),AC.CONST_DICTCOL_AMASS] = "organism_enterococcus_spp"
        df_dict.loc[df_dict[AC.CONST_DICTCOL_AMASS].str.contains("organism_salmonella"),AC.CONST_DICTCOL_AMASS] = "organism_salmonella_spp"
        dict_datavaltoamass = pd.Series(df_dict[AC.CONST_DICTCOL_AMASS].values,index=df_dict[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
        dict_amasstodataval = pd.Series(df_dict[AC.CONST_DICTCOL_DATAVAL].values,index=df_dict[AC.CONST_DICTCOL_AMASS].str.strip()).to_dict()
        #dict for value replacement in data column 
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["specimen_blood","specimen_cerebrospinal_fluid","specimen_genital_swab","specimen_respiratory_tract","specimen_stool","specimen_urine","specimen_others"])]
        dict_spectype_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["male","female"])]
        dict_gender_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS].isin(["community_origin","unknown_origin","hospital_origin"])]
        dict_inforg_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        temp_df = df_dict[df_dict[AC.CONST_DICTCOL_AMASS] == "died"]
        dict_died_datavaltoamass = pd.Series(temp_df[AC.CONST_DICTCOL_AMASS].values,index=temp_df[AC.CONST_DICTCOL_DATAVAL].astype("string").str.strip()).to_dict()
        #------------------------------------------------------------------------------------------------------------------------------------------------------
        # dictionary from AMASS_amr_const
        dict_ris = AC.dict_ris(df_dict)
        dict_ast = AC.dict_ast()
        dict_orgcatwithatb = AC.dict_orgcatwithatb(bisabom,bisentspp)
        dict_orgwithatb_mortality = AC.dict_orgwithatb_mortality(bisabom)
        dict_orgwithatb_incidence = AC.dict_orgwithatb_incidence(bisabom)
        list_antibiotic = AC.list_antibiotic
        AL.printlog("Complete process load data and dictionary: " + str(datetime.now()),False,logger)    
    except Exception as e: # work on python 3.x
        AL.printlog("Fail precoess load data and dictionary: " +  str(e),True,logger)
        logger.exception(e)
    #########################################################################################################################################
    # Start transform hosp and micro data
    try:
        # rename variable in dataval to amassval
        # What happen if key is duplicate
        AL.printlog("Note : Microbiology data column before trim/map with dictionary: " +  str(df_micro.columns),False,logger)
        df_micro.columns = df_micro.columns.str.strip()
        df_micro.rename(columns=dict_datavaltoamass, inplace=True)   
        AL.printlog("Note : Microbiology data column after trim/map with dictionary: " +  str(df_micro.columns),False,logger)
        if bishosp_ava:
            AL.printlog("Note : Hospital admission data column before trim/map with dictionary: " +  str(df_hosp.columns),False,logger)
            df_hosp.columns = df_hosp.columns.str.strip()
            df_hosp.rename(columns=dict_datavaltoamass, inplace=True)
            AL.printlog("Note : Hospital admission data column after trim/map with dictionary: " +  str(df_hosp.columns),False,logger)
        #CLean ward before remove column
        AL.printlog("-- Clean ward in micro --",False,logger) 
        fn_clean_ward(df_micro,AC.CONST_VARNAME_WARD,AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_NEWVARNAME_WARDTYPE,path_input,"dictionary_for_wards",logger) 
        #list all micro wards.
        try:
            df_micro_ward = df_micro.groupby(by=[AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_NEWVARNAME_WARDTYPE])[AC.CONST_NEWVARNAME_WARDCODE].count().reset_index(name='patient_count')
        except Exception as e:
            AL.printlog("Warning : Fail get micro ward list: " +  str(e),True,logger)
            logger.exception(e)
        try:
            temp_df = df_micro.groupby([AC.CONST_NEWVARNAME_WARDCODE, AC.CONST_VARNAME_WARD],dropna=False).size().reset_index(name='Count')
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_ward.xlsx", logger):
                print("Warning  : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_ward.xlsx")
        except Exception as e:
            AL.printlog("Warning : can't get list of ward in microbiology data file for display in data verification log "  + str(e),True,logger)
            logger.exception(e)
            pass
        #Remove unused column to save memory
        list_micorcol = [AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_SPECDATERAW,AC.CONST_VARNAME_COHO,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_SPECTYPE,AC.CONST_VARNAME_WARD,AC.CONST_NEWVARNAME_WARDCODE,AC.CONST_NEWVARNAME_WARDTYPE]
        list_micorcol = list_micorcol + list_antibiotic
        df_micro = AL.fn_keeponlycol(df_micro, list_micorcol)
        #-----------------------------------------------------------------------------------------------------------------------------------------    
        # Assign row number as unique key for each microbiology data row (This is not in R code)
        df_micro[AC.CONST_NEWVARNAME_MICROREC_ID] = np.arange(len(df_micro))
        #-----------------------------------------------------------------------------------------------------------------------------------------
        # Trim of space and unreadable charector for field that may need to map values such as spectype, organism.
        df_micro = AL.fn_df_strstrips(df_micro,[AC.CONST_VARNAME_SPECTYPE,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_SPECNUM,AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_COHO],logger)
        # Transform and map data
        df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE] = df_micro[AC.CONST_VARNAME_SPECTYPE].astype("string").map(dict_spectype_datavaltoamass)
        df_micro[AC.CONST_NEWVARNAME_BLOOD]  = df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE]
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] == "specimen_blood", AC.CONST_NEWVARNAME_BLOOD] = "blood"
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD] != "blood", AC.CONST_NEWVARNAME_BLOOD] = "non-blood"
        df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE] = df_micro[AC.CONST_NEWVARNAME_AMASSSPECTYPE].fillna("undefined")
        df_micro[AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_VARNAME_ORG].astype("string").map(dict_datavaltoamass).fillna(df_micro[AC.CONST_VARNAME_ORG])
        df_micro = clean_hn(df_micro,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_HN,logger)
        df_micro = fn_clean_date_andcalday_year_month(df_micro, AC.CONST_VARNAME_SPECDATERAW, AC.CONST_NEWVARNAME_CLEANSPECDATE, AC.CONST_NEWVARNAME_DAYTOSPECDATE, AC.CONST_NEWVARNAME_SPECYEAR, AC.CONST_NEWVARNAME_SPECMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
        df_micro = fn_clean_date_andcalday_year_month(df_micro, AC.CONST_VARNAME_SPECRPTDATERAW, AC.CONST_NEWVARNAME_CLEANSPECRPTDATE, AC.CONST_NEWVARNAME_DAYTOSPECRPTDATE, AC.CONST_NEWVARNAME_SPECRPTYEAR, AC.CONST_NEWVARNAME_SPECRPTMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
        
        dict_progvar["date_include_min"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].min().strftime("%d %b %Y")
        dict_progvar["date_include_max"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].max().strftime("%d %b %Y")
        # Transform data hm
        if bishosp_ava:
            AL.printlog("-- Clean ward in hosp --",False,logger) 
            fn_clean_ward(df_hosp,AC.CONST_VARNAME_WARD_HOSP,AC.CONST_NEWVARNAME_WARDCODE_HOSP,AC.CONST_NEWVARNAME_WARDTYPE_HOSP,path_input,"dictionary_for_wards",logger) 
            df_hosp = AL.fn_keeponlycol(df_hosp, [AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_ADMISSIONDATE,AC.CONST_VARNAME_DISCHARGEDATE,
                                                  AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_BIRTHDAY,AC.CONST_VARNAME_AGEY,AC.CONST_VARNAME_AGEGROUP,
                                                  AC.CONST_VARNAME_WARD_HOSP,AC.CONST_NEWVARNAME_WARDCODE_HOSP,AC.CONST_NEWVARNAME_WARDTYPE_HOSP])
            #Save log ward data verification log
            try:
                
                temp_df = df_hosp.groupby([AC.CONST_NEWVARNAME_WARDCODE_HOSP, AC.CONST_VARNAME_WARD_HOSP],dropna=False).size().reset_index(name='Count')
                if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_ward_hosp.xlsx", logger):
                    print("Warning  : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_ward_hosp.xlsx")
            except Exception as e:
                AL.printlog("Warning : can't get list of ward in hospital data file for display in data verification log "  + str(e),True,logger)
                logger.exception(e)
                pass
            # Trim of space and unreadable charector for field that may need to map values such as spectype, organism.
            df_hosp = AL.fn_df_strstrips(df_hosp,[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_COHO],logger)
            df_hosp = clean_hn(df_hosp,AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_HN_HOSP,logger)
            # Convert field to category type to save mem space
            df_hosp = AL.fn_df_tocategory_datatype(df_hosp,[AC.CONST_VARNAME_DISCHARGESTATUS,AC.CONST_VARNAME_GENDER,AC.CONST_VARNAME_AGEGROUP,AC.CONST_VARNAME_COHO],logger)
            #Build 3027 move ageyear to right place for calculate after clean
            # df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = 0
            # if dict_amasstodataval['age_year_available'].lower() == "yes":
            #     df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_VARNAME_AGEY].apply(pd.to_numeric,errors='coerce')
            # elif dict_amasstodataval['birthday_available'].lower() == "yes":
            #     df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] =  (df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE])/365.25
            # df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].apply(np.floor,errors='coerce')
            df_hosp_formerge = df_hosp.copy(deep=True)
            df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP] = df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].astype("string").map(dict_died_datavaltoamass).fillna(AC.CONST_ALIVE_VALUE) # From R code line 1154
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_ADMISSIONDATE, AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_DAYTOADMDATE, AC.CONST_NEWVARNAME_ADMYEAR, AC.CONST_NEWVARNAME_ADMMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_DISCHARGEDATE, AC.CONST_NEWVARNAME_CLEANDISDATE, AC.CONST_NEWVARNAME_DAYTODISDATE, AC.CONST_NEWVARNAME_DISYEAR, AC.CONST_NEWVARNAME_DISMONTHNAME, AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            df_hosp = fn_clean_date_andcalday_year_month(df_hosp, AC.CONST_VARNAME_BIRTHDAY, AC.CONST_NEWVARNAME_CLEANBIRTHDATE, AC.CONST_NEWVARNAME_DAYTOBIRTHDATE, "","", AC.CONST_CDATEFORMAT, AC.CONST_ORIGIN_DATE,logger)
            #Build 3027 move ageyear calculate after clean
            df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = 0
            try:
                if dict_amasstodataval['age_year_available'].lower() == "yes":
                    df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_VARNAME_AGEY].apply(pd.to_numeric,errors='coerce')
                elif dict_amasstodataval['birthday_available'].lower() == "yes":
                    df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] =  (df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOBIRTHDATE])/365.25
                df_hosp[AC.CONST_NEWVARNAME_AGEYEAR] = df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].apply(np.floor,errors='coerce')
            except:
                AL.printlog("Warning : unable to calculate age year ",False,logger)
            #V3.0.3 Calculate max date include
            df_hosp[AC.CONST_NEWVARNAME_DAYTOSTARTDATE] = df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE]
            df_hosp[AC.CONST_NEWVARNAME_DAYTOENDDATE] = df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE]
            try:
                dict_progvar["date_include_min"] = calcolmindayinclude(df_micro,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_hosp,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOSTARTDATE,AC.CONST_ORIGIN_DATE,dict_progvar["date_include_min"],logger)   
                dict_progvar["date_include_max"] = calcolmaxdayinclude(df_micro,AC.CONST_NEWVARNAME_CLEANSPECDATE,df_hosp,AC.CONST_NEWVARNAME_CLEANADMDATE,AC.CONST_NEWVARNAME_DAYTOENDDATE,AC.CONST_ORIGIN_DATE,dict_progvar["date_include_max"],logger) 
            except Exception as e:
                AL.printlog("Warning : calculate min max include date "  + str(e),True,logger)
                logger.exception(e)
                pass
            # Patient day
            df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY] = df_hosp[AC.CONST_NEWVARNAME_DAYTOENDDATE] - df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE] + 1
            df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO] = df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY] - 2
            df_hosp.loc[df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO] <0, AC.CONST_NEWVARNAME_PATIENTDAY_HO] = 0
            #debug_savecsv(df_hosp,path_repwithPID + "translated_hospital_data.csv",bisdebug,1,logger)
            AL.fn_savecsvwithencoding(df_hosp,path_repwithPID + "translated_hospital_data.csv",'utf-8-sig',1,logger)
        # Start Org Cat vs AST of interest -------------------------------------------------------------------------------------------------------
        # Suggest to be in a configuration files hide from user is better in term of coding
        df_micro["Temp" + AC.CONST_NEWVARNAME_ORG3] = df_micro[AC.CONST_NEWVARNAME_ORG3]
        df_micro.loc[df_micro[AC.CONST_NEWVARNAME_ORG3].astype("string").str.strip() == "","Temp" + AC.CONST_NEWVARNAME_ORG3] = AC.CONST_ORG_NOGROWTH
        df_micro[AC.CONST_NEWVARNAME_ORGCAT] = df_micro["Temp" + AC.CONST_NEWVARNAME_ORG3].map({k : vs[0] for k, vs in dict_orgcatwithatb.items()}).fillna(0)
        df_micro.drop("Temp" + AC.CONST_NEWVARNAME_ORG3, axis=1, inplace=True)              
        #Build 3027 add log for count mapped S\I\R to compare with logfile_ast for unmap values count
        temp_ris_mapped = pd.DataFrame(columns =["Antibiotics","frequency_raw"])      
        print(dict_ris)
        # Gen RIS columns
        for satb in list_antibiotic:
            if satb in df_micro.columns:
                df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb]  = ""
                try:
                    for s_toreplace in dict_ris:
                        #df_micro.loc[df_micro[satb].astype("string").str.contains(s_toreplace,regex=False), AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = dict_ris[s_toreplace] #v3.1 3032
                        try:
                            df_micro.loc[df_micro[satb].astype("string").str.strip().str.lower() == s_toreplace.lower(),
                                                 AC.CONST_NEWVARNAME_PREFIX_RIS + satb
                                                 ] = dict_ris[s_toreplace] #Change in 3.1 3101
                        except Exception as e:
                            AL.printlog("Waring: Unable to set " + AC.CONST_NEWVARNAME_PREFIX_RIS + satb + " columns base on " + satb + "column, this may because column cannot converted to string for comparing. " + str(e),False,logger)
                        #if still got those that not translated.
                        df_micro.loc[(df_micro[satb].astype("string").str.contains(s_toreplace,regex=False)) & (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] == ""), AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = dict_ris[s_toreplace] 
                    df_micro[satb] = df_micro[satb].astype("category")
                except Exception as e:
                    AL.printlog("Waring: Unable to set " + AC.CONST_NEWVARNAME_PREFIX_RIS + satb + " columns base on " + satb + "column, this may because original columns is duplicate during map with dictionary. " + str(e),False,logger)  
                    try:
                        if AC.CONST_ATBCOLDUP_ACTIONMODE == 1:
                            if isinstance(df_micro[satb], pd.DataFrame):
                                AL.printlog("Warning: confirm that columns is duplicated. Change the way to set data for "   + AC.CONST_NEWVARNAME_PREFIX_RIS + satb + "column. ",False,logger) 
                                for s_toreplace in dict_ris:
                                    for iii in range(len(df_micro[satb].columns)):
                                        temp_col = df_micro[satb].iloc[:,iii]
                                        #df_micro.loc[(temp_col.astype("string").str.contains(s_toreplace,regex=False) & (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] == "")), AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = dict_ris[s_toreplace] #v3.1 3032
                                        df_micro.loc[
                                                        (temp_col.astype("string").str.strip().str.lower() == s_toreplace.lower()) &
                                                        (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] == ""),
                                                        AC.CONST_NEWVARNAME_PREFIX_RIS + satb
                                                    ] = dict_ris[s_toreplace] #Change in 3.1 3101
                                        df_micro.loc[(temp_col.astype("string").str.contains(s_toreplace,regex=False) & (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] == "")), AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = dict_ris[s_toreplace]
                    except Exception as e:
                        AL.printlog("Warning : Still unable to set data for "   + AC.CONST_NEWVARNAME_PREFIX_RIS + satb + "column. " + str(e),False,logger)
                        logger.exception(e)
                try:
                    n_row = len(df_micro[(df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb].isnull() == False) & (df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] != "")])
                    onew_row = {"Antibiotics":satb,"frequency_raw":str(n_row)}     
                    temp_ris_mapped = pd.concat([temp_ris_mapped,pd.DataFrame([onew_row])], ignore_index = True)
                except Exception as e:
                    AL.printlog("Warning : unable to count value successfully map S\I\R for "   + AC.CONST_NEWVARNAME_PREFIX_RIS + satb + "column. " + str(e),False,logger)
                    logger.exception(e)
            else:
                df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = "" 
            df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb].astype("category")
        if not AL.fn_savexlsx(temp_ris_mapped, AC.CONST_PATH_RESULT + "logfile_ris_count.xlsx", logger):
            AL.printlog("Warning : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_ris_count.xlsx",False,logger)  
        del temp_ris_mapped
        # Gen AST columns
        #AL.printlog(dict_ast,False,logger) #v3.1 3032
        for satb in list_antibiotic:
            AL.printlog(AC.CONST_NEWVARNAME_PREFIX_RIS + satb,False,logger) #v3.1 3032
            df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_RIS + satb].map(dict_ast).fillna("NA") 
            df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb] = df_micro[AC.CONST_NEWVARNAME_PREFIX_AST + satb].astype("category")
        # Antibiotic group
        dict_antibioticgroup = AC.dict_atbgroup()
        df_micro = AL.fn_addatbgroupbyconfig(df_micro,dict_antibioticgroup,dict_ast,logger) 
        #3GCCBPN - RIS
        list_ast_asvalue = ["NR","R"]
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS] != "R") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS] != "R") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS].isna() == True)),
                     (df_micro[AC.CONST_NEWVARNAME_AST3GC_RIS] == "R") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS] != "R") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN_RIS].isna() == True))]
        df_micro[AC.CONST_NEWVARNAME_AST3GCCBPN_RIS] = np.select(temp_cond,list_ast_asvalue,default="-")
        #3GCCBPN -AST (NS)
        list_ast_asvalue = ["1","2"]
        temp_cond = [(df_micro[AC.CONST_NEWVARNAME_AST3GC] == "0") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True)),
                     (df_micro[AC.CONST_NEWVARNAME_AST3GC] == "1") & ((df_micro[AC.CONST_NEWVARNAME_ASTCBPN] != "1") | (df_micro[AC.CONST_NEWVARNAME_ASTCBPN].isna() == True))]
        df_micro[AC.CONST_NEWVARNAME_AST3GCCBPN] = np.select(temp_cond,list_ast_asvalue,default="0")
        
        list_amr_atb = AC.getlist_amr_atb(dict_orgcatwithatb)
        print(list_amr_atb)
        list_ast_atb = AC.getlist_ast_atb(dict_orgcatwithatb)
        print(list_ast_atb)
        if len(list_amr_atb) <=0:
            AL.printlog("Waring: list of antibiotic of interest for amr (get from dict_orgcatwithatb defined in AMASS_amr_const.py) is empty, application may pickup wrong record during deduplication by organism",False,logger)  
            df_micro[AC.CONST_NEWVARNAME_AMR] = 0
            df_micro[AC.CONST_NEWVARNAME_AMR_TESTED] = 0
            #Add in 3.1 to calculate number of R, I, and S as well as AST tested for each specimen
            df_micro[AC.CONST_NEWVARNAME_AST_R] = 0
            df_micro[AC.CONST_NEWVARNAME_AST_I] = 0
            df_micro[AC.CONST_NEWVARNAME_AST_S] = 0
            df_micro[AC.CONST_NEWVARNAME_AST_TESTED] = 0
        else:
            df_micro[AC.CONST_NEWVARNAME_AMR] = df_micro[list_amr_atb].apply(pd.to_numeric,errors='coerce').sum(axis=1, skipna=True)   
            df_micro[AC.CONST_NEWVARNAME_AMR_TESTED] = df_micro[list_amr_atb].apply(pd.to_numeric,errors='coerce').count(axis=1,numeric_only=True)   
            #Add in 3.1 to calculate number of R, I, and S as well as AST tested for each specimen
            df_micro[AC.CONST_NEWVARNAME_AST_R] = df_micro[list_ast_atb].apply(lambda row: (row == 'R').sum(), axis=1)
            df_micro[AC.CONST_NEWVARNAME_AST_I] = df_micro[list_ast_atb].apply(lambda row: (row == 'I').sum(), axis=1)
            df_micro[AC.CONST_NEWVARNAME_AST_S] = df_micro[list_ast_atb].apply(lambda row: (row == 'S').sum(), axis=1)
            df_micro[AC.CONST_NEWVARNAME_AST_TESTED] = df_micro[AC.CONST_NEWVARNAME_AST_R] + df_micro[AC.CONST_NEWVARNAME_AST_I] + df_micro[AC.CONST_NEWVARNAME_AST_S]
        
        #Micro remove and alter column data type to save memory usage --------------------------------------------------------------------------------------
        #Change text/object field type to category type to save memory usage
        df_micro = AL.fn_df_tocategory_datatype(df_micro,
                                                [AC.CONST_NEWVARNAME_BLOOD,AC.CONST_NEWVARNAME_ORG3,AC.CONST_NEWVARNAME_ORGCAT,AC.CONST_VARNAME_SPECTYPE,AC.CONST_VARNAME_ORG,AC.CONST_VARNAME_COHO,AC.CONST_NEWVARNAME_AMASSSPECTYPE],
                                                logger)
        ## Only blood specimen 
        df_micro_blood = df_micro.loc[df_micro[AC.CONST_NEWVARNAME_BLOOD]=="blood"]
        ## Only interesting org cat 
        df_micro_bsi = df_micro_blood.loc[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT]!=AC.CONST_ORG_NOTINTEREST_ORGCAT]
        #debug_savecsv(df_micro,path_repwithPID + "translated_microbiology.csv",bisdebug,1,logger)
        #debug_savecsv(df_micro_blood,path_repwithPID + "translated_microbiology_blood.csv",bisdebug,1,logger)
        #debug_savecsv(df_micro_bsi,path_repwithPID + "translated_microbiology_under_survey.csv",bisdebug,1,logger)
        AL.fn_savecsvwithencoding(df_micro,path_repwithPID + "translated_microbiology.csv",'utf-8-sig',1,logger)
        AL.fn_savecsvwithencoding(df_micro_blood,path_repwithPID + "translated_microbiology_blood.csv",'utf-8-sig',1,logger)
        AL.fn_savecsvwithencoding(df_micro_bsi,path_repwithPID + "translated_microbiology_under_survey.csv",'utf-8-sig',1,logger)
        #Version 3.0.3 Ward variables
        AL.printlog("Note : Microbiology data column when finish prepare data: " +  str(df_micro.columns),False,logger)
        if bishosp_ava:
            AL.printlog("Note : Hospital admission data column when finish prepare data: " +  str(df_hosp.columns),False,logger)
        AL.printlog("Complete prepare data: " + str(datetime.now()),False,logger)        
    except Exception as e: # work on python 3.x
        AL.printlog("Fail prepare data: " +  str(e),True,logger)   
        logger.exception(e)
    sub_printprocmem("before start merge",logger)
    # Start Merge with hosp data -------------------------------------------------------------------------------------------------------------------------
    df_datalog_mergedlist = pd.DataFrame()
    try:            
        #Get number of days allow specimen date to be before or after admnission period
        #list_specdate_tolerance =[AL.fn_getdict(dict_amasstodataval,AC.CONST_VARNAME_SPECDATE_BEFORE,"0"),AL.fn_getdict(dict_amasstodataval,AC.CONST_VARNAME_SPECDATE_AFTER,"0")]
        list_specdate_tolerance = [0,0]
        try:
            config=AL.readxlsxorcsv(AC.CONST_PATH_ROOT + "Configuration/", "Configuration",logger)
            config_lst = df_config.iloc[:,0].tolist()
            list_specdate_tolerance[0] = df_config.loc[config_lst.index("flexible_days_before_admission_for_CO"),"Setting parameters"]
            list_specdate_tolerance[1] =df_config.loc[config_lst.index("flexibile_days_after_discharge_for_HO"),"Setting parameters"]
        except:
            pass
        AL.printlog("Note : flexible days before and after admission period: " + str(list_specdate_tolerance) ,False,logger)   
        #Get from configuration instead
        df_hospmicro = fn_mergededup_hospmicro(df_micro, df_hosp_formerge, bishosp_ava,df_dict,dict_datavaltoamass,dict_inforg_datavaltoamass,dict_gender_datavaltoamass,dict_died_datavaltoamass,logger,df_datalog_mergedlist,list_specdate_tolerance)
        sub_printprocmem("finish merge micro and hosp data",logger)
        #Change to filter from df_hospmicro instead of call merge again
        df_hospmicro_blood = df_hospmicro.loc[df_hospmicro[AC.CONST_NEWVARNAME_BLOOD]=="blood"]
        sub_printprocmem("finish merge micro (blood) and hosp data",logger)
        df_hospmicro_bsi = df_hospmicro_blood.loc[df_hospmicro_blood[AC.CONST_NEWVARNAME_ORGCAT]!=AC.CONST_ORG_NOTINTEREST_ORGCAT]
        sub_printprocmem("finish merge micro (bsi) and hosp data",logger)
        del df_hosp_formerge
        #debug_savecsv(df_hospmicro,path_repwithPID + "merged_hospital_microbiology_COHO.csv",bisdebug,1,logger)
        #debug_savecsv(df_hospmicro_blood,path_repwithPID + "merged_hospital_microbiology_COHO_blood.csv",bisdebug,1,logger)
        #debug_savecsv(df_hospmicro_bsi,path_repwithPID + "merged_hospital_microbiology_COHO_under_survey.csv",bisdebug,1,logger)
        AL.fn_savecsvwithencoding(df_hospmicro,path_repwithPID + "merged_hospital_microbiology_COHO.csv",'utf-8-sig',1,logger)
        AL.fn_savecsvwithencoding(df_hospmicro_blood,path_repwithPID + "merged_hospital_microbiology_COHO_blood.csv",'utf-8-sig',1,logger)
        AL.fn_savecsvwithencoding(df_hospmicro_bsi,path_repwithPID + "merged_hospital_microbiology_COHO_under_survey.csv",'utf-8-sig',1,logger)
        gc.collect() 
        AL.printlog("Complete merge hosp and micro data: " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail merge hosp and micro data: " +  str(e),True,logger) 
        logger.exception(e)
    #########################################################################################################################################
    # Start summarize data - This part will not alter or add column to orignal micro data and merge hospital data 
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # General summary variable for calculation/report
    try:
        if bishosp_ava:
            dict_progvar["patientdays"] =int(df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY].sum())
            dict_progvar["patientdays_ho"] =int(df_hosp[AC.CONST_NEWVARNAME_PATIENTDAY_HO].sum())
            dict_progvar["hosp_date_min"] = df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].min().strftime("%d %b %Y")
            dict_progvar["hosp_date_max"] = df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].max().strftime("%d %b %Y")
        else:
            dict_progvar["patientdays"] ="NA"
            dict_progvar["patientdays_ho"] ="NA"
            dict_progvar["hosp_date_min"] = "NA"
            dict_progvar["hosp_date_max"] = "NA" 
        dict_progvar["micro_date_min"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].min().strftime("%d %b %Y")
        dict_progvar["micro_date_max"] = df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].max().strftime("%d %b %Y")
        dict_progvar["n_blood"] = len(df_micro_blood)
        dict_progvar["n_blood_pos"] = len(df_micro_blood[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT] != AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_blood_neg"] = len(df_micro_blood[df_micro_blood[AC.CONST_NEWVARNAME_ORGCAT] == AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_bsi_pos"] = len(df_micro_bsi[df_micro_bsi[AC.CONST_NEWVARNAME_ORGCAT] != AC.CONST_ORG_NOGROWTH_ORGCAT])
        dict_progvar["n_blood_patients"] = len(df_micro_blood[AC.CONST_NEWVARNAME_HN].unique())
        dict_progvar["checkpoint_section6"] = 0
        try:
            dict_progvar["checkpoint_section6"] = len(df_hospmicro_bsi[(df_hospmicro_bsi[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE) | (df_hospmicro_bsi[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE)])
        except:
            dict_progvar["checkpoint_section6"] = 0
        df_isoRep_blood = pd.DataFrame(columns=["Organism","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*",
                                                "Resistant(N)","Intermediate(N)",
                                                "Resistant(%)","Resistant-lower95CI(%)*","Resistant-upper95CI(%)*",
                                                "Intermediate(%)","Intermediate-lower95CI(%)*","Intermediate-upper95CI(%)*"])
        df_isoRep_blood_byorg = pd.DataFrame(columns=["Organism","Number_of_blood_specimens_culture_positive_for_the_organism"])
        df_isoRep_blood_byorg_dedup = pd.DataFrame(columns=["Organism","Number_of_blood_specimens_culture_positive_deduplicated"])
        df_isoRep_blood_incidence = pd.DataFrame(columns=["Organism","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"])
        df_isoRep_blood_incidence_atb = pd.DataFrame(columns=["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci","IncludeonlyR"])
        for sorgkey in dict_orgcatwithatb.keys():
            ocurorg = dict_orgcatwithatb[sorgkey]
            # skip no growth
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                temp_df = fn_deduplicatebyorgcat(df_micro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                # Start for atb sensitivity summary
                sorgname = ocurorg[2]
                df_isoRep_blood_byorg = pd.concat([df_isoRep_blood_byorg,pd.DataFrame([{"Organism":sorgname,"Number_of_blood_specimens_culture_positive_for_the_organism":len(df_micro_bsi[df_micro_bsi[AC.CONST_NEWVARNAME_ORGCAT]==int(scurorgcat)])}])],ignore_index = True)
                df_isoRep_blood_byorg_dedup = pd.concat([df_isoRep_blood_byorg_dedup,pd.DataFrame([{"Organism":sorgname,"Number_of_blood_specimens_culture_positive_deduplicated":len(temp_df)}])],ignore_index = True)
                list_atbname = ocurorg[3]
                list_atbmicrocol = ocurorg[4]
                # ADD CHECK if len list not the same avoid error
                for i in range(len(list_atbname)):
                    scuratbname = list_atbname[i]
                    scuratbmicrocol =list_atbmicrocol[i]
                    iRIS_S = 0
                    iRIS_I = 0
                    iRIS_R = 0
                    if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                        iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "R"])
                        iRIS_I = len(temp_df[temp_df[scuratbmicrocol] == "I"])
                        iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "S"])
                        
                    else:
                        iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "1"])
                        iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "0"])
                    itotal = iRIS_R +  iRIS_I + iRIS_S
                    iSuscep = iRIS_S
                    iNonsuscep = iRIS_R +  iRIS_I
                    nNonSuscepPercent = "NaN"
                    nLowerCI = "NA"
                    nUpperCI = "NA"
                    #Version 3.0.2
                    nR_percent = "NaN"
                    nI_percent = "NaN"
                    nR_LowerCI = "NA"
                    nR_UpperCI = "NA"
                    nI_LowerCI = "NA"
                    nI_UpperCI = "NA"
                    if itotal != 0 :
                        #nNonSuscepPercent = round(iNonsuscep / itotal, 2) * 100
                        nNonSuscepPercent = round(iNonsuscep / itotal*100, 1)
                        nLowerCI = fn_wilson_lowerCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                        nUpperCI  = fn_wilson_upperCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                        if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                            #nR_percent = round(iRIS_R / itotal, 2) * 100
                            nR_percent = round(iRIS_R / itotal*100, 1)
                            nR_LowerCI = fn_wilson_lowerCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                            nR_UpperCI  = fn_wilson_upperCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                            #nI_percent = round(iRIS_I / itotal, 2) * 100
                            nI_percent = round(iRIS_I / itotal*100, 1)
                            nI_LowerCI = fn_wilson_lowerCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                            nI_UpperCI  = fn_wilson_upperCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                    onew_row = {"Organism":sorgname,"Antibiotic":scuratbname,"Susceptible(N)":iSuscep,"Non-susceptible(N)": iNonsuscep,"Total(N)":itotal,
                                "Non-susceptible(%)":nNonSuscepPercent,"lower95CI(%)*":nLowerCI,"upper95CI(%)*" : nUpperCI,
                                "Resistant(N)":iRIS_R,"Intermediate(N)":iRIS_I,
                                "Resistant(%)":nR_percent,"Resistant-lower95CI(%)*":nR_LowerCI,"Resistant-upper95CI(%)*":nR_UpperCI,
                                "Intermediate(%)":nI_percent,"Intermediate-lower95CI(%)*":nI_LowerCI,"Intermediate-upper95CI(%)*":nI_UpperCI}   
                    df_isoRep_blood = pd.concat([df_isoRep_blood,pd.DataFrame([onew_row])],ignore_index = True)
                # Start incidence
                if sorgkey in dict_orgwithatb_incidence.keys():
                    ocurorg_incidence = dict_orgwithatb_incidence[sorgkey]
                    sorgname_incidence = ocurorg_incidence[0]
                    list_atbname_incidence = ocurorg_incidence[1]
                    list_atbmicrocol_incidence = ocurorg_incidence[2]
                    list_astvalue_incidence = ocurorg_incidence[3]
                    list_mode = ocurorg_incidence[4]
                    nPatient = len(temp_df)
                    nPercent = "NA"
                    nLowerCI = "NA"
                    nUpperCI = "NA"
                    if int(dict_progvar["n_blood_patients"]) > 0:
                        nPercent = (nPatient/dict_progvar["n_blood_patients"])*AC.CONST_PERPOP
                        nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                        nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                    onew_row = {"Organism":sorgname_incidence,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI  }   
                    df_isoRep_blood_incidence = pd.concat([df_isoRep_blood_incidence,pd.DataFrame([onew_row])],ignore_index = True)
                    for i in range(len(list_atbname_incidence)):
                        scuratbname = list_atbname_incidence[i]
                        scuratbmicrocol =list_atbmicrocol_incidence[i]
                        sCurastvalue = list_astvalue_incidence[i]
                        sMode = list_mode[i]
                        nPatient = 0
                        if type(sCurastvalue) == list:
                            nPatient = len(temp_df[temp_df[scuratbmicrocol].isin(sCurastvalue)])
                        else:
                            nPatient = len(temp_df[temp_df[scuratbmicrocol] == sCurastvalue])
                        nPercent = "NA"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        if int(dict_progvar["n_blood_patients"]) > 0:
                            nPercent = (nPatient/dict_progvar["n_blood_patients"])*AC.CONST_PERPOP
                            nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                            nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=dict_progvar["n_blood_patients"], conflevel=0.95, decimalplace=10)
                        onew_row = {"Organism":sorgname_incidence,"Priority_pathogen":scuratbname,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI,"IncludeonlyR":sMode  }   
                        #df_isoRep_blood_incidence_atb = df_isoRep_blood_incidence_atb.append(onew_row,ignore_index = True)
                        df_isoRep_blood_incidence_atb = pd.concat([df_isoRep_blood_incidence_atb,pd.DataFrame([onew_row])],ignore_index = True)
        AL.printlog("Complete analysis (By Sample base (Section 1,2,4,5,6)): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Sample base (Section 1,2,4,5,6)): " +  str(e),True,logger)     
        logger.exception(e)
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    # Summary data from hospmicro_blood and hospmicro_bsi 
    # Separate CO/HO dataframe
    try:
        #temp_df = df_hospmicro_blood.filter([AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_COHO_FINAL], axis=1)
        temp_df = df_hospmicro_blood[[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_CLEANSPECDATE,AC.CONST_NEWVARNAME_COHO_FINAL]].copy(deep=True)
        temp_df = fn_deduplicatedata(temp_df,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE],"first")
        temp_df2 = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_COHO_FINAL] == 0]
        temp_df = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_COHO_FINAL] == 1]
        dict_progvar["n_CO_blood_patients"] = len(temp_df2[AC.CONST_NEWVARNAME_HN].unique())
        dict_progvar["n_HO_blood_patients"] = len(temp_df[AC.CONST_NEWVARNAME_HN].unique())
        temp_df = temp_df.merge(temp_df2, how="inner", left_on=AC.CONST_NEWVARNAME_HN, right_on=AC.CONST_NEWVARNAME_HN,suffixes=("", "CO"))
        dict_progvar["n_2adm_firstbothCOHO_patients"] = len(temp_df[AC.CONST_NEWVARNAME_HN].unique())
    except Exception as e: # work on python 3.x
        AL.printlog("Fail summary for section 3, number of CO, HO patients: " +  str(e),True,logger)     
        logger.exception(e)
    try:
        df_COHO_isoRep_blood = pd.DataFrame(columns=["Organism","Infection_origin","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*",
                                                     "Resistant(N)","Intermediate(N)",
                                                     "Resistant(%)","Resistant-lower95CI(%)*","Resistant-upper95CI(%)*",
                                                     "Intermediate(%)","Intermediate-lower95CI(%)*","Intermediate-upper95CI(%)*"])

        df_COHO_isoRep_blood_byorg = pd.DataFrame(columns=["Organism","Number_of_patients_with_blood_culture_positive","Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file","Community_origin","Hospital_origin","Unknown_origin"])
        df_COHO_isoRep_blood_mortality = pd.DataFrame(columns=["Organism","Antibiotic","Infection_origin","Mortality","Mortality_lower_95ci","Mortality_upper_95ci","Number_of_deaths","Total_number_of_patients","IncludeonlyR"])
        df_COHO_isoRep_blood_mortality_byorg = pd.DataFrame(columns=["Organism","Infection_origin","Number_of_deaths","Total_number_of_patients"])
        df_COHO_isoRep_blood_incidence = pd.DataFrame(columns=["Organism","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci","Infection_origin"])
        df_COHO_isoRep_blood_incidence_atb = pd.DataFrame(columns=["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci","Infection_origin","IncludeonlyR"])
        for sorgkey in dict_orgcatwithatb.keys():
            ocurorg = dict_orgcatwithatb[sorgkey]
            # skip no growth
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                temp_df_byorgcat = fn_deduplicatebyorgcat_hospmico(df_hospmicro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                #debug_savecsv(temp_df_byorgcat,path_repwithPID + "section3_dedup_"  + ocurorg[2] + "_COHO.csv",bisdebug,1,logger)
                AL.fn_savecsvwithencoding(temp_df_byorgcat,path_repwithPID + "section3_dedup_"  + ocurorg[2] + "_COHO.csv",'utf-8-sig',1,logger)
                temp_df = fn_deduplicatebyorgcat(df_micro_bsi,AC.CONST_NEWVARNAME_ORGCAT, int(scurorgcat))
                iAll = len(temp_df)
                temp_df = fn_getunknownbyorg(ocurorg[2],temp_df_byorgcat,temp_df,AC.CONST_NEWVARNAME_HN,logger)
                #debug_savecsv(temp_df,path_repwithPID + "section3_dedup_"  + ocurorg[2] + "_unknown_origin.csv",bisdebug,1,logger)
                AL.fn_savecsvwithencoding(temp_df,path_repwithPID + "section3_dedup_"  + ocurorg[2] + "_unknown_origin.csv",'utf-8-sig',1,logger)
                iCO = 0
                iHO = 0
                for iCOHO in range(2):
                    # Dedup here may be differ from R code as R dedup before filter !!!
                    if iCOHO == 0:
                        sCOHO = AC.CONST_EXPORT_COHO_CO_DATAVAL
                        sCOHO_mortality = AC.CONST_EXPORT_COHO_MORTALITY_CO_DATAVAL
                        temp_df = temp_df_byorgcat.loc[temp_df_byorgcat[AC.CONST_NEWVARNAME_COHO_FINAL] == 0]
                        iCO = len(temp_df)
                    else:
                        sCOHO = AC.CONST_EXPORT_COHO_HO_DATAVAL
                        sCOHO_mortality = AC.CONST_EXPORT_COHO_MORTALITY_HO_DATAVAL
                        temp_df = temp_df_byorgcat.loc[temp_df_byorgcat[AC.CONST_NEWVARNAME_COHO_FINAL] == 1]
                        iHO = len(temp_df)
                    # Start for atb sensitivity summary                    
                    sorgname = ocurorg[2]
                    list_atbname = ocurorg[3]
                    list_atbmicrocol = ocurorg[4]
                    # ADD CHECK if len list not the same avoid error
                    for i in range(len(list_atbname)):
                        scuratbname = list_atbname[i]
                        scuratbmicrocol =list_atbmicrocol[i]
                        iRIS_S = 0
                        iRIS_I = 0
                        iRIS_R = 0
                        if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                            iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "R"])
                            iRIS_I = len(temp_df[temp_df[scuratbmicrocol] == "I"])
                            iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "S"])
                            
                        else:
                            iRIS_R = len(temp_df[temp_df[scuratbmicrocol] == "1"])
                            iRIS_S = len(temp_df[temp_df[scuratbmicrocol] == "0"])
                        itotal = iRIS_R +  iRIS_I + iRIS_S
                        iSuscep = iRIS_S
                        iNonsuscep = iRIS_R +  iRIS_I
                        nNonSuscepPercent = "NaN"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        #Version 3.0.2
                        nR_percent = "NaN"
                        nI_percent = "NaN"
                        nR_LowerCI = "NA"
                        nR_UpperCI = "NA"
                        nI_LowerCI = "NA"
                        nI_UpperCI = "NA"
                        if itotal != 0 :
                            #nNonSuscepPercent = round(iNonsuscep / itotal, 2) * 100
                            nNonSuscepPercent = round(iNonsuscep / itotal*100, 1)
                            nLowerCI = fn_wilson_lowerCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                            nUpperCI  = fn_wilson_upperCI(x=iNonsuscep, n=itotal, conflevel=0.95, decimalplace=1)
                            if scuratbmicrocol[0:len(AC.CONST_NEWVARNAME_PREFIX_RIS)] == AC.CONST_NEWVARNAME_PREFIX_RIS:
                                #nR_percent = round(iRIS_R / itotal, 2) * 100
                                nR_percent = round(iRIS_R / itotal*100, 1)
                                nR_LowerCI = fn_wilson_lowerCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                                nR_UpperCI  = fn_wilson_upperCI(x=iRIS_R, n=itotal, conflevel=0.95, decimalplace=1)
                                #nI_percent = round(iRIS_I / itotal, 2) * 100
                                nI_percent = round(iRIS_I / itotal*100, 1)
                                nI_LowerCI = fn_wilson_lowerCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                                nI_UpperCI  = fn_wilson_upperCI(x=iRIS_I, n=itotal, conflevel=0.95, decimalplace=1)
                        onew_row = {"Organism":sorgname,"Infection_origin":sCOHO,"Antibiotic":scuratbname,"Susceptible(N)":iSuscep,"Non-susceptible(N)": iNonsuscep,"Total(N)":itotal,
                                    "Non-susceptible(%)":nNonSuscepPercent,"lower95CI(%)*":nLowerCI,"upper95CI(%)*" : nUpperCI,
                                    "Resistant(N)":iRIS_R,"Intermediate(N)":iRIS_I,
                                    "Resistant(%)":nR_percent,"Resistant-lower95CI(%)*":nR_LowerCI,"Resistant-upper95CI(%)*":nR_UpperCI,
                                    "Intermediate(%)":nI_percent,"Intermediate-lower95CI(%)*":nI_LowerCI,"Intermediate-upper95CI(%)*":nI_UpperCI }
                        df_COHO_isoRep_blood = pd.concat([df_COHO_isoRep_blood,pd.DataFrame([onew_row])],ignore_index = True)  
                    # Start mortality
                    if sorgkey in dict_orgwithatb_mortality.keys():
                        ocurorg_mortality = dict_orgwithatb_mortality[sorgkey]
                        sorgname_mortality = ocurorg_mortality[0]
                        iOrgDied = len(temp_df[temp_df[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                        iOrgAlive = len(temp_df[temp_df[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                        iOrgTotal = iOrgDied + iOrgAlive
                        list_atbname_mortality = ocurorg_mortality[1]
                        list_atbmicrocol_mortality = ocurorg_mortality[2]
                        list_astvalue_mortality = ocurorg_mortality[3]
                        list_mode = ocurorg_mortality[4]
                        # ADD CHECK if len list not the same avoid error
                        for i in range(len(list_atbname_mortality)):
                            scuratbname = list_atbname_mortality[i]
                            scuratbmicrocol =list_atbmicrocol_mortality[i]
                            sCurastvalue = list_astvalue_mortality[i]
                            sMode = list_mode[i]
                            if type(sCurastvalue) == list:
                                temp_df2 = temp_df.loc[temp_df[scuratbmicrocol].isin(sCurastvalue)]
                            else:
                                temp_df2 = temp_df.loc[temp_df[scuratbmicrocol] == sCurastvalue]
                            iDied = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                            iAlive= len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                            itotal = iDied + iAlive
                            nLowerCI = 0
                            nUpperCI = 0
                            nDiedPercent = 0
                            sMortality = "NA"
                            if itotal != 0 :
                                nLowerCI = fn_wilson_lowerCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                                nUpperCI  = fn_wilson_upperCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                                nDiedPercent = int(round((iDied/itotal)*100, 0))
                                sMortality = str(nDiedPercent) + "% " + "(" + str(iDied) + "/" + str(itotal) + ")"
                            onew_row = {"Organism":sorgname_mortality,"Antibiotic":scuratbname,"Infection_origin":sCOHO_mortality,"Mortality":sMortality,
                                        "Mortality_lower_95ci":nLowerCI,"Mortality_upper_95ci":nUpperCI,"Number_of_deaths":iDied,"Total_number_of_patients":itotal,"IncludeonlyR":sMode}   
                            df_COHO_isoRep_blood_mortality = pd.concat([df_COHO_isoRep_blood_mortality,pd.DataFrame([onew_row])],ignore_index = True)         
                        df_COHO_isoRep_blood_mortality_byorg = pd.concat([df_COHO_isoRep_blood_mortality_byorg,pd.DataFrame([{"Organism":sorgname_mortality,"Infection_origin":sCOHO_mortality,
                                                                                                            "Number_of_deaths":iOrgDied,"Total_number_of_patients":iOrgTotal}])],ignore_index = True)
                    # Start incidence
                    if sorgkey in dict_orgwithatb_incidence.keys():
                        ocurorg_incidence = dict_orgwithatb_incidence[sorgkey]
                        sorgname_incidence = ocurorg_incidence[0]
                        list_atbname_incidence = ocurorg_incidence[1]
                        list_atbmicrocol_incidence = ocurorg_incidence[2]
                        list_astvalue_incidence = ocurorg_incidence[3]
                        list_mode = ocurorg_incidence[4]
                        nPatient = len(temp_df)
                        nPercent = "NA"
                        nLowerCI = "NA"
                        nUpperCI = "NA"
                        if iCOHO == 0: 
                            nTotal = dict_progvar["n_CO_blood_patients"]
                        else:
                            nTotal = dict_progvar["n_HO_blood_patients"]
                        if int(nTotal) > 0:
                            nPercent = (nPatient/nTotal)*AC.CONST_PERPOP
                            nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                            nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                        onew_row = {"Organism":sorgname_incidence,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI,"Infection_origin":sCOHO  }   
                        df_COHO_isoRep_blood_incidence = pd.concat([df_COHO_isoRep_blood_incidence,pd.DataFrame([onew_row])],ignore_index = True)
                        for i in range(len(list_atbname_incidence)):
                            scuratbname = list_atbname_incidence[i]
                            scuratbmicrocol =list_atbmicrocol_incidence[i]
                            sCurastvalue = list_astvalue_incidence[i]
                            sMode = list_mode[i]
                            #nPatient = len(temp_df[temp_df[scuratbmicrocol] == sCurastvalue])
                            nPatient = 0
                            if type(sCurastvalue) == list:
                                nPatient = len(temp_df[temp_df[scuratbmicrocol].isin(sCurastvalue)])
                            else:
                                nPatient = len(temp_df[temp_df[scuratbmicrocol] == sCurastvalue])
                            nPercent = "NA"
                            nLowerCI = "NA"
                            nUpperCI = "NA"
                            if int(nTotal) > 0:
                                nPercent = (nPatient/nTotal)*AC.CONST_PERPOP
                                nLowerCI = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                                nUpperCI  = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=nPatient, n=nTotal, conflevel=0.95, decimalplace=10)
                            onew_row = {"Organism":sorgname_incidence,"Priority_pathogen":scuratbname,"Number_of_patients":nPatient,"frequency_per_tested":nPercent,"frequency_per_tested_lci":nLowerCI,"frequency_per_tested_uci":nUpperCI,"Infection_origin":sCOHO,"IncludeonlyR":sMode  }   
                            df_COHO_isoRep_blood_incidence_atb = pd.concat([df_COHO_isoRep_blood_incidence_atb,pd.DataFrame([onew_row])],ignore_index = True) 
                #Save count CO/HO by org
                df_COHO_isoRep_blood_byorg = pd.concat([df_COHO_isoRep_blood_byorg,pd.DataFrame([{"Organism":sorgname,
                                                                                "Number_of_patients_with_blood_culture_positive":iAll,
                                                                                "Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file":iCO + iHO,
                                                                                "Community_origin":iCO,
                                                                                "Hospital_origin":iHO,
                                                                                "Unknown_origin":iAll- iCO- iHO
                                                                                }])],ignore_index = True)
        AL.printlog("Complete analysis (By Infection Origin (Section 3, 5, 6)): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Infection Origin (Section 3, 5, 6)): " +  str(e),True,logger)  
        logger.exception(e)
    sub_printprocmem("finish section 1-6",logger)
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    # ANNEX A
    try:
        dict_annex_a_listorg = AC.dict_annex_a_listorg
        dict_annex_a_spectype = AC.dict_annex_a_spectype
        
        df_dict_annex_a = df_dict_micro.copy(deep=True)
        #Reverse order of dict -> to be the same as R when map
        df_dict_annex_a = df_dict_annex_a[::-1].reset_index(drop = True) 
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_brucella"),AC.CONST_DICTCOL_AMASS] = "organism_brucella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_shigella"),AC.CONST_DICTCOL_AMASS] = "organism_shigella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_vibrio"),AC.CONST_DICTCOL_AMASS] = "organism_vibrio_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "organism_salmonella_typhi",AC.CONST_DICTCOL_AMASS] = "t_typhi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "organism_salmonella_paratyphi",AC.CONST_DICTCOL_AMASS] = "t_paratyphi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS].str.contains("organism_salmonella"),AC.CONST_DICTCOL_AMASS] = "organism_non-typhoidal_salmonella_spp"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "t_typhi",AC.CONST_DICTCOL_AMASS] = "organism_salmonella_typhi"
        df_dict_annex_a.loc[df_dict_annex_a[AC.CONST_DICTCOL_AMASS] == "t_paratyphi",AC.CONST_DICTCOL_AMASS] = "organism_salmonella_paratyphi"
        dict_datavaltoamass_annex_a = pd.Series(df_dict_annex_a[AC.CONST_DICTCOL_AMASS].values,index=df_dict_annex_a[AC.CONST_DICTCOL_DATAVAL].str.strip()).to_dict()
        df_micro_annex_a = df_micro.copy(deep=True)
        df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA] =  df_micro_annex_a[AC.CONST_VARNAME_SPECTYPE].astype("string").str.strip().map(dict_datavaltoamass_annex_a).fillna("specimen_others")
        df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA] = df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA].str.strip().map(dict_annex_a_spectype).fillna("Others")
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA] =  df_micro_annex_a[AC.CONST_VARNAME_ORG].astype("string").str.strip().map(dict_datavaltoamass_annex_a).fillna(AC.CONST_ANNEXA_NON_ORG)
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORGCAT_ANNEXA] =  df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[0] for k, vs in dict_annex_a_listorg.items()}).fillna(0)
        df_micro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA] =  df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[2] for k, vs in dict_annex_a_listorg.items()}).fillna("")
        df_micro_annex_a = df_micro_annex_a.loc[df_micro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=""]
        df_micro_annex_a = df_micro_annex_a.loc[df_micro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA]!=AC.CONST_ANNEXA_NON_ORG]
        # Start Annex A summary table -----------------------------------------------------------------------------------------------------------
        temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                     ['microbiology_data','Number_of_all_culture_positive', len(df_micro_annex_a)]
                     ]
        for sspeckey in dict_annex_a_spectype:
            sspecname = dict_annex_a_spectype[sspeckey]
            sparameter = "Number_of_" + sspecname.replace("\n"," ").replace(" ","_").lower().strip() + "_culture_positive"
            n_annexa_speccount = len(df_micro_annex_a[df_micro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA]==sspecname])
            temp_list.append(['microbiology_data',sparameter,str(n_annexa_speccount)])
        df_annexA = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"])
        # Start Annex A1 summary table ----------------------------------------------------------------------------------------------------------
        temp_df = fn_deduplicatedata(df_micro_annex_a,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")
        temp_df2 = fn_deduplicatedata(df_micro_annex_a,[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True],"last",[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")

        df_annexA1_pivot = temp_df.pivot_table(index=AC.CONST_NEWVARNAME_ORGNAME_ANNEXA, columns=AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA, aggfunc={AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA:len}, fill_value=0)
        df_annexA1_pivot.columns = df_annexA1_pivot.columns.droplevel(0)
        # add missing column and row to pivot table
        df_annexA1 = pd.DataFrame(columns=["Organism","Total number\nof patients*"])
        for oorg in dict_annex_a_listorg.values():
            if oorg[1] == 1 :
                sorg = oorg[2]
                list_rowvalue = [sorg]
                for sspec in dict_annex_a_spectype.values():
                    # add new column on the fly
                    if sspec not in df_annexA1.columns:
                        df_annexA1[sspec] = 0
                    # if have value for this org and specimen type then use it value
                    iCur = 0
                    if sorg in df_annexA1_pivot.index:
                        if sspec in df_annexA1_pivot.columns:
                            iCur = df_annexA1_pivot.loc[sorg][sspec]
                    list_rowvalue.append(iCur)
                # append row  
                itotal = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]==sorg])
                list_rowvalue.insert(1,itotal)
                df_annexA1.loc[sorg] = list_rowvalue 
        df_annexA1.loc['Total'] = df_annexA1.sum()
        df_annexA1.loc['Total','Organism'] = "Total"
        #put NA back in
        for sspec in dict_annex_a_spectype.values():
            if sspec not in df_annexA1_pivot.columns:
                try:
                    df_annexA1[sspec] = "NA"
                    AL.printlog("Note : " + sspec + " set to NA for annex A1",False,logger)
                except:
                    pass
        sub_printprocmem("finish analyse annex A1 data",logger)
        # Start Annex A1b,A2 summary table ----------------------------------------------------------------------------------------------------------
        df_hospmicro_annex_a = df_hospmicro
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_VARNAME_SPECTYPE].astype("string").str.strip().map(dict_datavaltoamass_annex_a).fillna("specimen_others")
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA] = df_hospmicro_annex_a[AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA].str.strip().map(dict_annex_a_spectype).fillna("Others")
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_VARNAME_ORG].astype("string").str.strip().map(dict_datavaltoamass_annex_a).fillna(AC.CONST_ANNEXA_NON_ORG)
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGCAT_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[0] for k, vs in dict_annex_a_listorg.items()}).fillna(0)
        df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA] =  df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORG3_ANNEXA].map({k : vs[2] for k, vs in dict_annex_a_listorg.items()}).fillna("")
        df_hospmicro_annex_a = df_hospmicro_annex_a.loc[df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=""]
        df_hospmicro_annex_a = df_hospmicro_annex_a.loc[df_hospmicro_annex_a[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]!=AC.CONST_ANNEXA_NON_ORG]
        temp_df = fn_deduplicatedata(df_hospmicro_annex_a,[AC.CONST_VARNAME_HOSPITALNUMBER, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER],"first")
        #A1b ---------------------------------------------------------------------------
        #the total number per specimen
        #AL.fn_savecsv(df_hospmicro_annex_a, AC.CONST_PATH_ROOT + "df_hospmicro_annex_a.csv", 2, logger)
        #Need to deduplicate in case wrong admission data, 1 patient stay on multiple ward at the same time.
        temp_df_A11_sum = fn_deduplicatedata(df_hospmicro_annex_a,[AC.CONST_NEWVARNAME_MICROREC_ID],[True],"last",[AC.CONST_NEWVARNAME_MICROREC_ID],"first")
        temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]],
                     ['hospital_admission_data','Minimum_admdate', dict_progvar["hosp_date_min"] if bishosp_ava else "NA"], 
                     ['hospital_admission_data','Maximum_admdate', dict_progvar["hosp_date_max"] if bishosp_ava else "NA"], 
                     ['microbiology_data','Number_of_all_culture_positive', len(temp_df_A11_sum)]
                     ]
        
        for sspeckey in dict_annex_a_spectype:
            sspecname = dict_annex_a_spectype[sspeckey]
            sparameter = "Number_of_" + sspecname.replace("\n"," ").replace(" ","_").lower().strip() + "_culture_positive"
            n_annexa_speccount = len(temp_df_A11_sum[temp_df_A11_sum[AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA]==sspecname])
            temp_list.append(['microbiology_data',sparameter,str(n_annexa_speccount)])
        df_annexA_A11_Sum = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"])
        #the table
        temp_df_A11 = fn_deduplicatedata(df_hospmicro_annex_a,[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_SPECTYPE_ANNEXA,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")
        temp_df_A11_2 = fn_deduplicatedata(df_hospmicro_annex_a,[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA, AC.CONST_NEWVARNAME_CLEANSPECDATE],[True,True,True],"last",[AC.CONST_VARNAME_HOSPITALNUMBER,AC.CONST_NEWVARNAME_ORGCAT_ANNEXA],"first")
        df_annexA11_pivot = temp_df_A11.pivot_table(index=AC.CONST_NEWVARNAME_ORGNAME_ANNEXA, columns=AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA, aggfunc={AC.CONST_NEWVARNAME_SPECTYPENAME_ANNEXA:len}, fill_value=0)
        df_annexA11_pivot.columns = df_annexA11_pivot.columns.droplevel(0)
        # add missing column and row to pivot table
        df_annexA11 = pd.DataFrame(columns=["Organism","Total number\nof patients*"])
        for oorg in dict_annex_a_listorg.values():
            if oorg[1] == 1 :
                sorg = oorg[2]
                list_rowvalue = [sorg]
                for sspec in dict_annex_a_spectype.values():
                    # add new column on the fly
                    if sspec not in df_annexA11.columns:
                        df_annexA11[sspec] = 0
                    # if have value for this org and specimen type then use it value
                    iCur = 0
                    if sorg in df_annexA11_pivot.index:
                        if sspec in df_annexA11_pivot.columns:
                            iCur = df_annexA11_pivot.loc[sorg][sspec]
                    list_rowvalue.append(iCur)
                # append row  
                itotal = len(temp_df_A11_2[temp_df_A11_2[AC.CONST_NEWVARNAME_ORGNAME_ANNEXA]==sorg])
                list_rowvalue.insert(1,itotal)
                df_annexA11.loc[sorg] = list_rowvalue 
        df_annexA11.loc['Total'] = df_annexA11.sum()
        df_annexA11.loc['Total','Organism'] = "Total"
        #put NA back in
        for sspec in dict_annex_a_spectype.values():
            if sspec not in df_annexA11_pivot.columns:
                try:
                    df_annexA11[sspec] = "NA"
                    AL.printlog("Note : " + sspec + " set to NA for annex A1b",False,logger)
                except:
                    pass
        sub_printprocmem("finish analyse annex A1b data",logger)
        df_annexA2 = pd.DataFrame(columns=["Organism","Number_of_deaths","Total_number_of_patients","Mortality(%)","Mortality_lower_95ci","Mortality_upper_95ci"])
        for sorgkey in dict_annex_a_listorg:
            ocurorg = dict_annex_a_listorg[sorgkey]
            if ocurorg[1] == 1 :
                scurorgcat = ocurorg[0]
                sorgname = ocurorg[2]
                temp_df2 = temp_df.loc[temp_df[AC.CONST_NEWVARNAME_ORG3_ANNEXA]==sorgkey]
                iDied = len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_DIED_VALUE])
                iAlive= len(temp_df2[temp_df2[AC.CONST_NEWVARNAME_DISOUTCOME] == AC.CONST_ALIVE_VALUE])
                itotal = iDied + iAlive
                nLowerCI = 0
                nUpperCI = 0
                nDiedPercent = 0
                sMortality = "NA"
                if itotal != 0 :
                    nLowerCI = fn_wilson_lowerCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                    nUpperCI  = fn_wilson_upperCI(x=iDied, n=itotal, conflevel=0.95, decimalplace=1)
                    nDiedPercent = int(round((iDied/itotal)*100, 0))
                onew_row = {"Organism":sorgname,"Number_of_deaths":iDied,"Total_number_of_patients":itotal,
                            "Mortality(%)":nDiedPercent,"Mortality_lower_95ci":nLowerCI,"Mortality_upper_95ci":nUpperCI }     
                df_annexA2 = pd.concat([df_annexA2,pd.DataFrame([onew_row])],ignore_index = True)
        sub_printprocmem("finish analyse annex A2 data",logger)
        AL.printlog("Complete analysis (Annex A): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail analysis (By Infection Origin (Section 3, 5, 6)): " +  str(e),True,logger) 
        logger.exception(e)        
    
    #########################################################################################################################################
    # Start export data - in future may also directly call function for generate report here by passing dataframe for better performance
    # --------------------------------------------------------------------------------------------------------------------------------------------------
    try:
        # --------------------------------------------------------------------------------------------------------------------------------------------------             
        # SECTION 1 - Number
        df_report1_page3 = pd.DataFrame(columns=["Type_of_data_file","Parameters","Values"])
        temp_var_list = [["hospital_name","Hospital_name"],["country","Country"],["contact_person","Contact_person"],["contact_address","Contact_address"],["contact_email","Contact_email"],["notes_on_the_cover","notes_on_the_cover"]]
        for ovar in temp_var_list:
            temp_list = ["microbiology_data", ovar[1],  dict_amasstodataval[ovar[0]]]
            df_report1_page3.loc[len(df_report1_page3)] = temp_list
        if bishosp_ava:
            soverall_mindate = dict_progvar["date_include_min"]
            soverall_maxdate = dict_progvar["date_include_max"]
            try:
                soverall_mindate = soverall_mindate.strftime("%d %b %Y")
                soverall_maxdate = soverall_maxdate.strftime("%d %b %Y")
            except:
                pass
            temp_list = [['overall_data','Minimum_date', soverall_mindate], 
                         ['overall_data','Maximum_date', soverall_maxdate], 
                         ['microbiology_data','Number_of_records', len(df_micro)], 
                         ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                         ['hospital_admission_data','Number_of_records', len(df_hosp)], 
                         ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"]], 
                         ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"]],
                         ['hospital_admission_data','Patient_days', dict_progvar["patientdays"]],
                         ['hospital_admission_data','Patient_days_his', dict_progvar["patientdays_ho"]]]     
        else:
            temp_list = [['overall_data','Minimum_date', dict_progvar["date_include_min"]], 
                         ['overall_data','Maximum_date', dict_progvar["date_include_max"]], 
                         ['microbiology_data','Number_of_records', len(df_micro)], 
                         ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]]]
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        df_report1_page3 = pd.concat([df_report1_page3,temp_df],ignore_index = True)
        if not AL.fn_savecsv(df_report1_page3, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec1_res_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec1_res_i)
        # SECTION 1 - by month
        df_report1_page4 = pd.DataFrame(data={"Month":fn_allmonthname()},index=fn_allmonthname())
        temp_df2 = df_micro.groupby([AC.CONST_NEWVARNAME_SPECMONTHNAME])[AC.CONST_NEWVARNAME_HN].count().reset_index(name='Number_of_specimen_in_microbiology_data_file')
        df_report1_page4 = df_report1_page4.merge(temp_df2, how="left", left_on="Month", right_on=AC.CONST_NEWVARNAME_SPECMONTHNAME,suffixes=("", "_MICRO"))
        df_report1_page4.drop(AC.CONST_NEWVARNAME_SPECMONTHNAME, axis=1, inplace=True)
        df_report1_page4['Number_of_specimen_in_microbiology_data_file'] = df_report1_page4['Number_of_specimen_in_microbiology_data_file'].fillna(0)
        if bishosp_ava:
            temp_df2 = df_hosp.groupby([AC.CONST_NEWVARNAME_ADMMONTHNAME])[AC.CONST_NEWVARNAME_HN_HOSP].count().reset_index(name='Number_of_hospital_records_in_hospital_admission_data_file')
            df_report1_page4 = df_report1_page4.merge(temp_df2, how="left", left_on="Month", right_on=AC.CONST_NEWVARNAME_ADMMONTHNAME,suffixes=("", "_HOSP"))
            df_report1_page4.drop(AC.CONST_NEWVARNAME_ADMMONTHNAME, axis=1, inplace=True)
            df_report1_page4['Number_of_hospital_records_in_hospital_admission_data_file'] = df_report1_page4['Number_of_hospital_records_in_hospital_admission_data_file'].fillna(0)
        if not AL.fn_savecsv(df_report1_page4, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec1_num_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec1_num_i)
        #Generate V2 Compatible files in resultdata folder
        AL.printlog("Start generate section 1 result data file compatible with AMASS V2.0: " ,False,logger)
        try:
            if not AL.fn_savecsv(df_report1_page3[df_report1_page3["Type_of_data_file"] !='overall_data'], AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec1_res_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec1_res_i)
            if not AL.fn_savecsv(df_report1_page4, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec1_num_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec1_num_i)
        except Exception as e:
            AL.printlog("Fail generate section 1 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
            logger.exception(e) 
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 2 - Summary
        temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]],
                     ['microbiology_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                     ['microbiology_data','Number_of_blood_culture_negative', dict_progvar["n_blood_neg"]], 
                     ['microbiology_data','Number_of_blood_culture_positive', dict_progvar["n_blood_pos"]], 
                     ['microbiology_data','Number_of_blood_culture_positive_for_organism_under_this_survey', dict_progvar["n_bsi_pos"]]]
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_res_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_res_i)
        # SECTION 2 - By Org, Org dedup, AMR proportion
        if not AL.fn_savecsv(df_isoRep_blood, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_amr_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_amr_i)
        if not AL.fn_savecsv(df_isoRep_blood_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_org_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_org_i)
        if not AL.fn_savecsv(df_isoRep_blood_byorg_dedup, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_pat_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec2_pat_i)
        #Generate V2 Compatible files in resultdata folder
        AL.printlog("Start generate section 2 result data file compatible with AMASS V2.0: " ,False,logger)
        try:
            #table and graph
            v2_df_isoRep_blood = V2.fn_combine_orgdata(logger,df_isoRep_blood,
                                                             dict_combineorg=V2.CONST_COMBINE_ORG,
                                                             lst_groupbycol=['Antibiotic'],
                                                             lst_sumcol=["Susceptible(N)","Non-susceptible(N)","Total(N)"],
                                                             col_org="Organism",col_totaln="Total(N)",col_n="Non-susceptible(N)",col_percent="Non-susceptible(%)",col_lower95CI="lower95CI(%)*",col_upper95CI="upper95CI(%)*")
            v2_df_isoRep_blood = V2.fn_filterorg_atb(v2_df_isoRep_blood ,col_org="Organism",col_atb="Antibiotic", dict_orgatb=V2.CONST_V2_DICT_ORG_ATB)
            v2_df_isoRep_blood = v2_df_isoRep_blood[["Organism","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*"]]
            #change atbname
            for satb in V2.CONST_V2_LIST_RENAME_ATB:
                v2_df_isoRep_blood.loc[v2_df_isoRep_blood["Antibiotic"] == satb,"Antibiotic"] = V2.CONST_V2_LIST_RENAME_ATB[satb]
            #All before dedup
            v2_df_isoRep_blood_byorg = V2.fn_combine_orgdata(logger,df_isoRep_blood_byorg,
                                                             dict_combineorg=V2.CONST_COMBINE_ORG,
                                                             lst_groupbycol=[],
                                                             lst_sumcol=["Number_of_blood_specimens_culture_positive_for_the_organism"],
                                                             col_org="Organism",col_totaln="",col_n="",col_percent="",col_lower95CI="",col_upper95CI="")
            v2_df_isoRep_blood_byorg = V2.fn_filterorg(v2_df_isoRep_blood_byorg ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG)
            #Dedup
            v2_df_isoRep_blood_byorg_dedup = V2.fn_combine_orgdata(logger,df_isoRep_blood_byorg_dedup,
                                                             dict_combineorg=V2.CONST_COMBINE_ORG,
                                                             lst_groupbycol=[],
                                                             lst_sumcol=["Number_of_blood_specimens_culture_positive_deduplicated"],
                                                             col_org="Organism",col_totaln="",col_n="",col_percent="",col_lower95CI="",col_upper95CI="")
            v2_df_isoRep_blood_byorg_dedup = V2.fn_filterorg(v2_df_isoRep_blood_byorg_dedup ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG)
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_res_i, 2, logger):
               print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_res_i)
               # SECTION 2 - By Org, Org dedup, AMR proportion
            if not AL.fn_savecsv(v2_df_isoRep_blood, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_amr_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_amr_i)
            if not AL.fn_savecsv(v2_df_isoRep_blood_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_org_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_org_i)
            if not AL.fn_savecsv(v2_df_isoRep_blood_byorg_dedup, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_pat_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec2_pat_i)
        except Exception as e:
            AL.printlog("Fail generate section 2 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
            logger.exception(e)  
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 3 - Summary
        if bishosp_ava:
            temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]],  
                         ['merged_data','Number_of_patients_with_blood_culture_positive_for_organism_under_this_survey', df_COHO_isoRep_blood_byorg["Number_of_patients_with_blood_culture_positive"].sum()], 
                         ['merged_data','Number_of_patients_with_community_origin_BSI', df_COHO_isoRep_blood_byorg["Community_origin"].sum()], 
                         ['merged_data','Number_of_patients_with_hospital_origin_BSI', df_COHO_isoRep_blood_byorg["Hospital_origin"].sum()], 
                         ['merged_data','Number_of_patients_with_unknown_origin_BSI', df_COHO_isoRep_blood_byorg["Unknown_origin"].sum()]]
            temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_res_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_res_i)
            # SECTION 3 - By Org, Org dedup, AMR proportion
            if not AL.fn_savecsv(df_COHO_isoRep_blood, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_amr_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_amr_i)
            if not AL.fn_savecsv(df_COHO_isoRep_blood_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_pat_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec3_pat_i)
        #Generate V2 Compatible files in resultdata folder
        AL.printlog("Start generate section 3 result data file compatible with AMASS V2.0: " ,False,logger)
        try:
            #table and graph
            v2_df_COHO_isoRep_blood = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood,
                                                             dict_combineorg=V2.CONST_COMBINE_ORG,
                                                             lst_groupbycol=['Infection_origin','Antibiotic'],
                                                             lst_sumcol=["Susceptible(N)","Non-susceptible(N)","Total(N)"],
                                                             col_org="Organism",col_totaln="Total(N)",col_n="Non-susceptible(N)",col_percent="Non-susceptible(%)",col_lower95CI="lower95CI(%)*",col_upper95CI="upper95CI(%)*")
            v2_df_COHO_isoRep_blood = V2.fn_filterorg_atb(v2_df_COHO_isoRep_blood,col_org="Organism",col_atb="Antibiotic", dict_orgatb=V2.CONST_V2_DICT_ORG_ATB)
            v2_df_COHO_isoRep_blood = v2_df_COHO_isoRep_blood[["Organism","Infection_origin","Antibiotic","Susceptible(N)","Non-susceptible(N)","Total(N)","Non-susceptible(%)","lower95CI(%)*","upper95CI(%)*"]]
            v2_temp_df = v2_df_COHO_isoRep_blood[v2_df_COHO_isoRep_blood["Infection_origin"] != AC.CONST_EXPORT_COHO_CO_DATAVAL]
            v2_df_COHO_isoRep_blood = v2_df_COHO_isoRep_blood[v2_df_COHO_isoRep_blood["Infection_origin"] == AC.CONST_EXPORT_COHO_CO_DATAVAL]
            v2_df_COHO_isoRep_blood = pd.concat([v2_df_COHO_isoRep_blood,v2_temp_df],ignore_index = True)
            #change atbname
            for satb in V2.CONST_V2_LIST_RENAME_ATB:
                v2_df_COHO_isoRep_blood.loc[v2_df_COHO_isoRep_blood["Antibiotic"] == satb,"Antibiotic"] = V2.CONST_V2_LIST_RENAME_ATB[satb]
            #Summary page
            v2_df_COHO_isoRep_blood_byorg = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_byorg,
                                                             dict_combineorg=V2.CONST_COMBINE_ORG,
                                                             lst_groupbycol=[],
                                                             lst_sumcol=["Number_of_patients_with_blood_culture_positive","Number_of_patients_with_blood_culture_positive_merged_with_hospital_data_file","Community_origin","Hospital_origin","Unknown_origin"],
                                                             col_org="Organism",col_totaln="",col_n="",col_percent="",col_lower95CI="",col_upper95CI="")
            v2_df_COHO_isoRep_blood_byorg = V2.fn_filterorg(v2_df_COHO_isoRep_blood_byorg ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG)
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_res_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_res_i)
            # SECTION 3 - By Org, Org dedup, AMR proportion
            if not AL.fn_savecsv(v2_df_COHO_isoRep_blood, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_amr_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_amr_i)
            if not AL.fn_savecsv(v2_df_COHO_isoRep_blood_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_pat_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec3_pat_i)

        except Exception as e:
            AL.printlog("Fail generate section 3 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
            logger.exception(e)  
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 4 - Summary
        if dict_progvar["n_blood_neg"] > 0:
            temp_list = [['merged_data','Minimum_date', dict_progvar["micro_date_min"]], 
                         ['merged_data','Maximum_date', dict_progvar["micro_date_max"]],  
                         ['merged_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                         ['merged_data','Number_of_patients_sampled_for_blood_culture', dict_progvar["n_blood_patients"]]]
            temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
            if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_res_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_res_i)
            # SECTION 4 By Org, Org and atb
            if not AL.fn_savecsv(df_isoRep_blood_incidence, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_blo_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_blo_i)
            if not AL.fn_savecsv(df_isoRep_blood_incidence_atb, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_pri_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec4_pri_i)
            #Generate V2 Compatible files in resultdata folder
            AL.printlog("Start generate section 4 result data file compatible with AMASS V2.0: " ,False,logger)
            try:
                #table and graph
                v2_df_isoRep_blood_incidence = V2.fn_combine_orgdata(logger,df_isoRep_blood_incidence,
                                                                 dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                 lst_groupbycol=[],
                                                                 lst_sumcol=["Number_of_patients"],
                                                                 col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",sec4_5_totaln=dict_progvar["n_blood_patients"])
                v2_df_isoRep_blood_incidence = V2.fn_filterorg(v2_df_isoRep_blood_incidence ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5)                
                #All before dedup
                v2_df_isoRep_blood_incidence_atb = V2.fn_rename_beforecombine(logger,df_isoRep_blood_incidence_atb[df_isoRep_blood_incidence_atb["IncludeonlyR"]==0],"Priority_pathogen",V2.CONST_RENAME_SEC4_5)
                #v2_df_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,df_isoRep_blood_incidence_atb[df_isoRep_blood_incidence_atb["IncludeonlyR"]==0],
                v2_df_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,v2_df_isoRep_blood_incidence_atb,
                                                                 dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                 lst_groupbycol=["Priority_pathogen"],
                                                                 lst_sumcol=["Number_of_patients"],
                                                                 col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",sec4_5_totaln=dict_progvar["n_blood_patients"])
                v2_df_isoRep_blood_incidence_atb = V2.fn_filterorg(v2_df_isoRep_blood_incidence_atb ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5)
                v2_df_isoRep_blood_incidence_atb = v2_df_isoRep_blood_incidence_atb[["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci"]]
                if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_res_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_res_i)
                # SECTION 4 By Org, Org and atb
                if not AL.fn_savecsv(v2_df_isoRep_blood_incidence, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_blo_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_blo_i)
                if not AL.fn_savecsv(v2_df_isoRep_blood_incidence_atb, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_pri_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec4_pri_i)

            except Exception as e:
                AL.printlog("Fail generate section 4 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
                logger.exception(e)      
            
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 5 - Summary
        if bishosp_ava:
            if dict_progvar["n_blood_neg"] > 0:
                temp_df = df_hospmicro_blood[[AC.CONST_NEWVARNAME_HN,AC.CONST_NEWVARNAME_CLEANADMDATE]].dropna().drop_duplicates()
                temp_df = temp_df.groupby([AC.CONST_NEWVARNAME_HN])[AC.CONST_NEWVARNAME_CLEANADMDATE].count().reset_index(name="count")
                temp_list = [['merged_data','Minimum_date', dict_progvar["micro_date_min"]], 
                             ['merged_data','Maximum_date', dict_progvar["micro_date_max"]],  
                             ['merged_data','Number_of_blood_specimens_collected', dict_progvar["n_blood"]], 
                             ['merged_data','Number_of_patients_sampled_for_blood_culture', dict_progvar["n_blood_patients"]],
                             ['merged_data','Number_of_patients_with_blood_culture_within_first_2_days_of_admission', dict_progvar["n_CO_blood_patients"]], 
                             ['merged_data','Number_of_patients_with_blood_culture_within_after_2_days_of_admission', dict_progvar["n_HO_blood_patients"]],
                             ['merged_data','Number_of_patients_with_unknown_origin',dict_progvar["n_blood_patients"] - len(df_hospmicro_blood[AC.CONST_NEWVARNAME_HN].unique()) ],
                             ['merged_data','Number_of_patients_had_more_than_one_admission', dict_progvar["n_2adm_firstbothCOHO_patients"]]]
                temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_res_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_res_i)
                # SECTION 5 By Org, Org and atb by CO/HO
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL], AC.CONST_PATH_RESULT +AC.CONST_FILENAME_sec5_com_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_com_i)
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence_atb[df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL], AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_com_amr_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_com_amr_i)
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL], AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_hos_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_hos_i)
                if not AL.fn_savecsv(df_COHO_isoRep_blood_incidence_atb[df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL], AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_hos_amr_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec5_hos_amr_i)
                #Generate V2 Compatible files in resultdata folder
                AL.printlog("Start generate section 5 result data file compatible with AMASS V2.0: " ,False,logger)
                try:
                    #by org CO
                    v2_df_CO_isoRep_blood_incidence = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL],
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                     lst_groupbycol=["Infection_origin"],
                                                                     lst_sumcol=["Number_of_patients"],
                                                                     col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",
                                                                     sec4_5_totaln=dict_progvar["n_CO_blood_patients"])
                    v2_df_CO_isoRep_blood_incidence = V2.fn_filterorg(v2_df_CO_isoRep_blood_incidence ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5)  
                    #by org HO
                    v2_df_HO_isoRep_blood_incidence = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_incidence[df_COHO_isoRep_blood_incidence["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL],
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                     lst_groupbycol=["Infection_origin"],
                                                                     lst_sumcol=["Number_of_patients"],
                                                                     col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",
                                                                     sec4_5_totaln=dict_progvar["n_HO_blood_patients"])
                    v2_df_HO_isoRep_blood_incidence = V2.fn_filterorg(v2_df_HO_isoRep_blood_incidence ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5) 
                    #by pathogen CO
                    v2_df_CO_isoRep_blood_incidence_atb = V2.fn_rename_beforecombine(logger,df_COHO_isoRep_blood_incidence_atb[(df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL) & (df_COHO_isoRep_blood_incidence_atb["IncludeonlyR"]==0)],"Priority_pathogen",V2.CONST_RENAME_SEC4_5)
                    #v2_df_CO_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_incidence_atb[(df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_CO_DATAVAL) & (df_COHO_isoRep_blood_incidence_atb["IncludeonlyR"]==0)],
                    v2_df_CO_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,v2_df_CO_isoRep_blood_incidence_atb,
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                     lst_groupbycol=["Infection_origin","Priority_pathogen"],
                                                                     lst_sumcol=["Number_of_patients"],
                                                                     col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",
                                                                     sec4_5_totaln=dict_progvar["n_CO_blood_patients"])
                    v2_df_CO_isoRep_blood_incidence_atb = V2.fn_filterorg(v2_df_CO_isoRep_blood_incidence_atb ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5)
                    v2_df_CO_isoRep_blood_incidence_atb = v2_df_CO_isoRep_blood_incidence_atb[["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci","Infection_origin"]]
                    #by pathogen HO
                    v2_df_HO_isoRep_blood_incidence_atb = V2.fn_rename_beforecombine(logger,df_COHO_isoRep_blood_incidence_atb[(df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL) & (df_COHO_isoRep_blood_incidence_atb["IncludeonlyR"]==0)],"Priority_pathogen",V2.CONST_RENAME_SEC4_5)
                    #v2_df_HO_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_incidence_atb[(df_COHO_isoRep_blood_incidence_atb["Infection_origin"]==AC.CONST_EXPORT_COHO_HO_DATAVAL) & (df_COHO_isoRep_blood_incidence_atb["IncludeonlyR"]==0)],
                    v2_df_HO_isoRep_blood_incidence_atb = V2.fn_combine_orgdata(logger,v2_df_HO_isoRep_blood_incidence_atb,        
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG_SEC4_5,
                                                                     lst_groupbycol=["Infection_origin","Priority_pathogen"],
                                                                     lst_sumcol=["Number_of_patients"],
                                                                     col_org="Organism",col_totaln="",col_n="Number_of_patients",col_percent="frequency_per_tested",col_lower95CI="frequency_per_tested_lci",col_upper95CI="frequency_per_tested_uci",
                                                                     sec4_5_totaln=dict_progvar["n_HO_blood_patients"])
                    v2_df_HO_isoRep_blood_incidence_atb = V2.fn_filterorg(v2_df_HO_isoRep_blood_incidence_atb ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG_SEC4_5)
                    v2_df_HO_isoRep_blood_incidence_atb =  v2_df_HO_isoRep_blood_incidence_atb[["Organism","Priority_pathogen","Number_of_patients","frequency_per_tested","frequency_per_tested_lci","frequency_per_tested_uci","Infection_origin"]]
                    temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                    if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_res_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_res_i)
                    # SECTION 5 By Org, Org and atb by CO/HO
                    if not AL.fn_savecsv(v2_df_CO_isoRep_blood_incidence, AC.CONST_PATH_RESULT +AC.CONST_FILENAME_V2_sec5_com_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_com_i)
                    if not AL.fn_savecsv(v2_df_CO_isoRep_blood_incidence_atb, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_com_amr_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_com_amr_i)
                    if not AL.fn_savecsv(v2_df_HO_isoRep_blood_incidence, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_hos_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_hos_i)
                    if not AL.fn_savecsv(v2_df_HO_isoRep_blood_incidence_atb, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_hos_amr_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec5_hos_amr_i)
                except Exception as e:
                    AL.printlog("Fail generate section 5 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
                    logger.exception(e)      
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION 6 - Summary
        if bishosp_ava:
            if dict_progvar["checkpoint_section6"] > 0:
                n_hosp_unique = len(df_hosp[AC.CONST_NEWVARNAME_HN_HOSP].unique())
                #n_hosp_died_unique = len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP]==AC.CONST_DIED_VALUE][AC.CONST_NEWVARNAME_HN_HOSP].unique())
                n_hosp_died_unique = len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_DISOUTCOME_HOSP]==AC.CONST_DIED_VALUE])
                n_died_per = int(round((n_hosp_died_unique/n_hosp_unique)*100, 0))
                n_died_per_text = str(n_died_per) +"% (" + str(n_hosp_died_unique) + "/" + str(n_hosp_unique) + ")"
                temp_list = [['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                             ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                             ['merged_data','Number_of_blood_culture_positive_for_organism_under_this_survey', df_COHO_isoRep_blood_byorg["Number_of_patients_with_blood_culture_positive"].sum()], 
                             ['merged_data','Number_of_patients_with_community_origin_BSI', df_COHO_isoRep_blood_byorg["Community_origin"].sum()], 
                             ['merged_data','Number_of_patients_with_hospital_origin_BSI', df_COHO_isoRep_blood_byorg["Hospital_origin"].sum()],
                             ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"]], 
                             ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"]],
                             ['merged_data','Number_of_records', len(df_hosp) if bishosp_ava == True else 'NA'], 
                             ['merged_data','Number_of_patients_included', n_hosp_unique if bishosp_ava == True else 'NA'], 
                             ['merged_data','Number_of_deaths', n_hosp_died_unique if bishosp_ava == True else 'NA'], 
                             ['merged_data','Mortality', n_died_per_text if bishosp_ava == True else 'NA']
                             ]
                temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_res_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_res_i)
                # SECTION 6 By Org, Org and atb by CO/HO
                if not AL.fn_savecsv(df_COHO_isoRep_blood_mortality_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_mor_byorg_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_mor_byorg_i)
                if not AL.fn_savecsv(df_COHO_isoRep_blood_mortality, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_mor_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_sec6_mor_i)
                #Generate V2 Compatible files in resultdata folder
                AL.printlog("Start generate section 6 result data file compatible with AMASS V2.0: ",False,logger)
                try:
                    #table and graph
                    v2_df_COHO_isoRep_blood_mortality_byorg = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_mortality_byorg,
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG,
                                                                     lst_groupbycol=["Infection_origin"],
                                                                     lst_sumcol=["Number_of_deaths","Total_number_of_patients"],
                                                                     col_org="Organism")
                    v2_df_COHO_isoRep_blood_mortality_byorg = V2.fn_filterorg(v2_df_COHO_isoRep_blood_mortality_byorg ,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG)                
                    #All before dedup
                    v2_df_COHO_isoRep_blood_mortality = V2.fn_combine_orgdata(logger,df_COHO_isoRep_blood_mortality[df_COHO_isoRep_blood_mortality["IncludeonlyR"]==0],
                                                                     dict_combineorg=V2.CONST_COMBINE_ORG,
                                                                     lst_groupbycol=["Infection_origin","Antibiotic"],
                                                                     lst_sumcol=["Number_of_deaths","Total_number_of_patients"],
                                                                     col_org="Organism",col_totaln="Total_number_of_patients",col_n="Number_of_deaths",col_percent="Mortality",col_lower95CI="Mortality_lower_95ci",col_upper95CI="Mortality_upper_95ci",
                                                                     bIsSec6=True)
                    v2_df_COHO_isoRep_blood_mortality = V2.fn_filterorg(v2_df_COHO_isoRep_blood_mortality,col_org="Organism", lst_org=V2.CONST_V2_LIST_ORG)
                    v2_df_COHO_isoRep_blood_mortality = v2_df_COHO_isoRep_blood_mortality[["Organism","Antibiotic","Infection_origin","Mortality","Mortality_lower_95ci","Mortality_upper_95ci","Number_of_deaths","Total_number_of_patients"]]
                    v2_temp_df = v2_df_COHO_isoRep_blood_mortality[v2_df_COHO_isoRep_blood_mortality["Infection_origin"] != AC.CONST_EXPORT_COHO_MORTALITY_CO_DATAVAL]
                    v2_df_COHO_isoRep_blood_mortality = v2_df_COHO_isoRep_blood_mortality[v2_df_COHO_isoRep_blood_mortality["Infection_origin"] == AC.CONST_EXPORT_COHO_MORTALITY_CO_DATAVAL]
                    v2_df_COHO_isoRep_blood_mortality = pd.concat([v2_df_COHO_isoRep_blood_mortality,v2_temp_df],ignore_index = True)
                    temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
                    if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_res_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_res_i)
                    # SECTION 6 By Org, Org and atb by CO/HO
                    if not AL.fn_savecsv(v2_df_COHO_isoRep_blood_mortality_byorg, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_mor_byorg_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_mor_byorg_i)
                    if not AL.fn_savecsv(v2_df_COHO_isoRep_blood_mortality, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_mor_i, 2, logger):
                        print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_sec6_mor_i)
                except Exception as e:
                    AL.printlog("Fail generate section 6 result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
                    logger.exception(e)
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # SECTION ANNEX A - Summary
        if not AL.fn_savecsv(df_annexA, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_res_i , 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_res_i)
        # A1
        if not AL.fn_savecsv(df_annexA1, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_pat_i, 2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_pat_i)
        # A1b
        if bishosp_ava:
            if not AL.fn_savecsv(df_annexA_A11_Sum, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_res_i_A11 , 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_res_i_A11)
            if not AL.fn_savecsv(df_annexA11, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_pat_i_A11, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_pat_i_A11)
        # A2
        if bishosp_ava:
            if not AL.fn_savecsv(df_annexA2, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_mor_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_FILENAME_secA_mor_i)
        #Generate V2 Compatible files in resultdata folder
        AL.printlog("Start generate annex A result data file compatible with AMASS V2.0: ",False,logger)
        try:
            if not AL.fn_savecsv(df_annexA, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_secA_res_i , 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_V2_FILENAME_secA_res_i)
            # A1
            if not AL.fn_savecsv(df_annexA1, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_secA_pat_i, 2, logger):
                print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_V2_FILENAME_secA_pat_i)
            # A2
            if bishosp_ava:
                if not AL.fn_savecsv(df_annexA2, AC.CONST_PATH_RESULT + AC.CONST_FILENAME_V2_secA_mor_i, 2, logger):
                    print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + AC.CONST_V2_FILENAME_secA_mor_i)
        except Exception as e:
            AL.printlog("Fail generate annex A result data file compatible with AMASS V2.0: " +  str(e),True,logger) 
            logger.exception(e)
        # --------------------------------------------------------------------------------------------------------------------------------------------------
        # Data verification log
        #3.0.2
        ispecnohn = 0
        ispecnotadm = 0
        imergeall = 0
        imergeblood = 0
        imergebsi = 0
        try:
            temp_df=fn_notindata(df_micro, df_datalog_mergedlist, AC.CONST_NEWVARNAME_MICROREC_ID, AC.CONST_NEWVARNAME_MICROREC_ID,True)
            ispecnohn = len(temp_df)
            temp_df=fn_notindata(df_datalog_mergedlist,df_hospmicro, AC.CONST_NEWVARNAME_MICROREC_ID, AC.CONST_NEWVARNAME_MICROREC_ID,True)
            ispecnotadm = len(temp_df)
            imergeall = len(df_hospmicro[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
            imergeblood = len(df_hospmicro_blood[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
            imergebsi=len(df_hospmicro_bsi[AC.CONST_NEWVARNAME_MICROREC_ID].drop_duplicates())
        except:
            pass
        iHN_leadzero_count_micro = 0
        iHN_leadzero_count_hosp = 0
        try:
            iHN_leadzero_count_micro = len(df_micro[df_micro[AC.CONST_NEWVARNAME_HN].astype(str).str[0]=="0"])
            if bishosp_ava:
                iHN_leadzero_count_hosp = len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_HN_HOSP].astype(str).str[0]=="0"])
        except:
            pass
        AL.printlog("Complete export data for report section 1-6, Annex A: " + str(datetime.now()),False,logger)
        #Data logger
        #Build 3027
        swrongspecdate = "NA"
        swrongadmdate = "NA"
        swrongdisdate = "NA"
        try:
            AC.CONST_ORIGIN_DATE 
            d2000 = pd.to_datetime("2000-01-01")
            d2100 = pd.to_datetime("2100-12-31")
            day2000 = (d2000- AC.CONST_ORIGIN_DATE ).days
            day2100 = (d2100- AC.CONST_ORIGIN_DATE ).days
            swrongspecdate = str(len(df_micro[(df_micro[AC.CONST_NEWVARNAME_DAYTOSPECDATE].notnull()) & ((df_micro[AC.CONST_NEWVARNAME_DAYTOSPECDATE].fillna(0) < day2000) | (df_micro[AC.CONST_NEWVARNAME_DAYTOSPECDATE].fillna(0) > day2100))]))
            swrongadmdate = str(len(df_hosp[(df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE].notnull()) & ((df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE].fillna(0) < day2000) | (df_hosp[AC.CONST_NEWVARNAME_DAYTOADMDATE].fillna(0) > day2100))]))
            swrongdisdate = str(len(df_hosp[(df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE].notnull()) & ((df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE].fillna(0) < day2000) | (df_hosp[AC.CONST_NEWVARNAME_DAYTODISDATE].fillna(0) > day2100))]))
        except Exception as e:
            AL.printlog("Warning : Count wrong date/wrong date format error : " + str(e),True,logger)
            logger.exception(e)
        temp_list = [['microbiology_data','Number_of_records', len(df_micro)], 
                     ['microbiology_data','Minimum_date', dict_progvar["micro_date_min"]], 
                     ['microbiology_data','Maximum_date', dict_progvar["micro_date_max"]], 
                     ['microbiology_data','Number_of_missing_specimen_date',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECDATERAW].isnull()]))],
                     ['microbiology_data','Number_of_missing_specimen_type',str(len(df_micro[df_micro[AC.CONST_VARNAME_SPECTYPE].isnull()]))],
                     ['microbiology_data','Number_of_missing_culture_result',str(len(df_micro[df_micro[AC.CONST_VARNAME_ORG].isnull()]))],
                     ['hospital_admission_data','Number_of_records', len(df_hosp) if bishosp_ava else "NA"], 
                     ['hospital_admission_data','Minimum_date', dict_progvar["hosp_date_min"] if bishosp_ava else "NA"], 
                     ['hospital_admission_data','Maximum_date', dict_progvar["hosp_date_max"] if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_admission_date',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_ADMISSIONDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_discharge_date',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGEDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_outcome_result',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_DISCHARGESTATUS].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_age',str(len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_AGEYEAR].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_gender',str(len(df_hosp[df_hosp[AC.CONST_VARNAME_GENDER].isnull()])) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_unmatchhn',str(ispecnohn) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_unmatchperiod',str(ispecnotadm) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchall',str(imergeall) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchblood',str(imergeblood) if bishosp_ava else "NA"],
                     ['merged_data','Number_of_matchbsi',str( imergebsi) if bishosp_ava else "NA"],
                     ['microbiology_data','Hospital_name',dict_amasstodataval["hospital_name"]],
                     ['microbiology_data','Country',dict_amasstodataval["country"]],
                     ['microbiology_data','Number_of_HN_with_leadingzero',str(iHN_leadzero_count_micro)],
                     ['hospital_admission_data','Number_of_HN_with_leadingzero',str(iHN_leadzero_count_hosp)],
                     ['microbiology_data','Number_of_missing_or_unknown_specimen_date',str(len(df_micro[df_micro[AC.CONST_NEWVARNAME_CLEANSPECDATE].isnull()]))],
                     ['hospital_admission_data','Number_of_missing_or_unknown_admission_date',str(len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_CLEANADMDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_missing_or_unknown_discharge_date',str(len(df_hosp[df_hosp[AC.CONST_NEWVARNAME_CLEANDISDATE].isnull()])) if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_wrong_admission_date',swrongadmdate if bishosp_ava else "NA"],
                     ['hospital_admission_data','Number_of_wrong_discharge_date',swrongdisdate if bishosp_ava else "NA"],
                     ['microbiology_data','Number_of_wrong_specimen_date',swrongspecdate]
                     ]
        temp_df = pd.DataFrame(temp_list, columns =["Type_of_data_file","Parameters","Values"]) 
        if not AL.fn_savecsv(temp_df, AC.CONST_PATH_RESULT + "logfile_results.csv",2, logger):
            print("Error : Cannot save csv file : " + AC.CONST_PATH_RESULT + "logfile_results.csv")
        temp_df = df_micro.groupby([AC.CONST_VARNAME_ORG])[AC.CONST_VARNAME_ORG].count().reset_index(name='Frequency')
        temp_df.rename(columns={AC.CONST_VARNAME_ORG: 'Organism'}, inplace=True) 
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_organism.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_organism.xlsx")
        temp_df = df_micro.groupby([AC.CONST_VARNAME_SPECTYPE])[AC.CONST_VARNAME_SPECTYPE].count().reset_index(name='Frequency')
        temp_df.rename(columns={AC.CONST_VARNAME_SPECTYPE: 'Specimen'}, inplace=True)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_specimen.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_specimen.xlsx")
        temp_df = check_date_format(df_micro, AC.CONST_VARNAME_SPECDATERAW, logger)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatspecdate.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatspecdate.xlsx")
        temp_df = check_hn_format(df_micro, AC.CONST_NEWVARNAME_HN, logger)
        if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_microhn.xlsx", logger):
            print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_microhn.xlsx")
        if bishosp_ava:
            temp_df = df_hosp.groupby([AC.CONST_VARNAME_GENDER])[AC.CONST_VARNAME_GENDER].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_VARNAME_GENDER: 'Gender'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_gender.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_gender.xlsx")
            temp_df = df_hosp.groupby([AC.CONST_VARNAME_DISCHARGESTATUS])[AC.CONST_VARNAME_DISCHARGESTATUS].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_VARNAME_DISCHARGESTATUS: 'Discharge status'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_discharge.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_discharge.xlsx")
            temp_df = df_hosp.groupby([AC.CONST_NEWVARNAME_AGEYEAR])[AC.CONST_NEWVARNAME_AGEYEAR].count().reset_index(name='Frequency')
            temp_df.rename(columns={AC.CONST_NEWVARNAME_AGEYEAR: 'Age'}, inplace=True)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_age.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_age.xlsx")
            temp_df = check_date_format(df_hosp, AC.CONST_VARNAME_ADMISSIONDATE, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatadmdate.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatadmdate.xlsx")
            temp_df = check_date_format(df_hosp, AC.CONST_VARNAME_DISCHARGEDATE, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_formatdisdate.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_formatdisdate.xlsx")
            temp_df = check_hn_format(df_hosp, AC.CONST_NEWVARNAME_HN_HOSP, logger)
            if not AL.fn_savexlsx(temp_df, AC.CONST_PATH_RESULT + "logfile_hosphn.xlsx", logger):
                print("Error : Cannot save xlsx file : " + AC.CONST_PATH_RESULT + "logfile_hosphn.xlsx")
        AL.printlog("Complete Export data for data verifcation log (END analysis): " + str(datetime.now()),False,logger)
    except Exception as e: # work on python 3.x
        AL.printlog("Fail export data section 1-6, Appendix A: " +  str(e),True,logger)  
        logger.exception(e)
    AL.printlog("End AMR analysis: " + str(datetime.now()),False,logger)
    sub_printprocmem("finish AMR analysis",logger)
    ANNEX_B.generate_annex_b(df_dict_micro, df_micro_annexb,logger,AC.CONST_ANNEXB_USING_MAPPEDDATA,False)
    sub_printprocmem("finish Analysis ANNEX B",logger)
    if bIsDoAnnexC == True:
        if (len(df_hospmicro) > 0) & (len(df_hospmicro_blood) > 0):
            try:
                import AMASS_annex_c_analysis as ANNEX_C
                import AMASS_annex_c_const as ACC
                bisrunannexc = True
                if (ACC.CONST_RUN_ANNEXC_WITH_NOHOSP == True) | (bishosp_ava == True):
                    bisrunannexc = True
                else:
                    bisrunannexc = False
                if bisrunannexc == True:
                    ANNEX_C.prepare_fromHospMicro_toSaTScan(logger,df_all=df_hospmicro, df_blo=df_hospmicro_blood)
                    ANNEX_C.call_SaTScan(logger,path_output=AC.CONST_PATH_TEMPWITH_PID,prmfile=ACC.CONST_FILENAME_NEWPARAM)
                    ANNEX_C.prepare_annexc_results(logger,b_wardhighpat=True, num_wardhighpat=2)
                    sub_printprocmem("finish Analysis ANNEX C",logger)
                else:
                    sub_printprocmem("Not run analysis ANNEX C set to run only with hosp data and no hosp available",logger)
            except Exception as e:
                AL.printlog("Error : Import/run analysis ANNEX C : " + str(e),True,logger)
                logger.exception(e)
        else:
            sub_printprocmem("No hospmicrodata for analysis ANNEX C",logger)
    else:
        sub_printprocmem("Command line specify not do Annex C",logger)
    AMR_REPORT_NEW.generate_amr_report(df_dict_micro,dict_orgcatwithatb,dict_orgwithatb_mortality,dict_orgwithatb_incidence,df_micro_ward,bishosp_ava,logger)
    sub_printprocmem("finish generate report",logger)
    SUP_REPORT.generate_supplementary_report(df_dict_micro,logger,AC.CONST_ANNEXB_USING_MAPPEDDATA)
    sub_printprocmem("finish generate suppelmentary report",logger)
    
# ------------------------------------------------------------------------------------------------------------------------------------------------------      
#Main loop
mainloop()