# -*- coding: utf-8 -*-
"""
Propose as common function for transform AMASS V3.0 analysed data in each section to file compatible with version 2 resultdata.
Created on Thu Nov  9 12:35:12 2023
@author: Prapass
"""
import AMASS_amr_commonlib as AL
import AMASS_amr_const as AC
import pandas as pd
from scipy.stats import norm

CONST_COMBINE_ORG = {AC.CONST_ORG_ENTEROCOCCUS_SPP:[AC.CONST_ORG_ENTEROCOCCUS_FAECALIS,AC.CONST_ORG_ENTEROCOCCUS_FAECIUM]}
CONST_COMBINE_ORG_SEC4_5 = {AC.CONST_ORG_ENTEROCOCCUS_SPP:["E. faecalis","E. faecium"]}
CONST_RENAME_SEC4_5 = {"Vancomycin-NSE. faecalis":"Vancomycin-NSEnterococcus spp.","Vancomycin-NSE. faecium":"Vancomycin-NSEnterococcus spp."}
CONST_V2_LIST_ORG = [AC.CONST_ORG_STAPHYLOCOCCUS_AUREUS, AC.CONST_ORG_ENTEROCOCCUS_SPP, AC.CONST_ORG_STREPTOCOCCUS_PNEUMONIAE, 
              AC.CONST_ORG_SALMONELLA_SPP, AC.CONST_ORG_ESCHERICHIA_COLI, AC.CONST_ORG_KLEBSIELLA_PNEUMONIAE, 
              AC.CONST_ORG_PSEUDOMONAS_AERUGINOSA, AC.CONST_ORG_ACINETOBACTER_BAUMANNII]
CONST_V2_LIST_ORG_SEC4_5 = ['S. aureus', AC.CONST_ORG_ENTEROCOCCUS_SPP, 'S. pneumoniae', AC.CONST_ORG_SALMONELLA_SPP, 'E. coli', 'K. pneumoniae', 'P. aeruginosa', 'A. baumannii']
CONST_V2_DICT_ORG_ATB = {AC.CONST_ORG_STAPHYLOCOCCUS_AUREUS:['Methicillin', 'Vancomycin','Clindamycin'],
                         AC.CONST_ORG_ENTEROCOCCUS_SPP:['Ampicillin', 'Vancomycin','Teicoplanin','Linezolid','Daptomycin'], 
                         AC.CONST_ORG_STREPTOCOCCUS_PNEUMONIAE:['Penicillin G', 'Oxacillin','Co-trimoxazole','3GC','Ceftriaxone','Cefotaxime','Erythromycin','Clindamycin','Levofloxacin'], 
                         AC.CONST_ORG_SALMONELLA_SPP:['FLUOROQUINOLONES', 'Ciprofloxacin','Levofloxacin','3GC','Ceftriaxone','Cefotaxime','Ceftazidime','CARBAPENEMS','Imipenem','Meropenem','Ertapenem','Doripenem'],
                         AC.CONST_ORG_ESCHERICHIA_COLI:['Gentamicin', 'Amikacin','Co-trimoxazole','Ampicillin','FLUOROQUINOLONES','Ciprofloxacin','Levofloxacin','3GC','Cefpodoxime','Ceftriaxone','Cefotaxime','Ceftazidime','Cefepime','CARBAPENEMS','Imipenem','Meropenem','Ertapenem','Doripenem','Colistin'],
                         AC.CONST_ORG_KLEBSIELLA_PNEUMONIAE:['Gentamicin', 'Amikacin','Co-trimoxazole','FLUOROQUINOLONES','Ciprofloxacin','Levofloxacin','3GC','Cefpodoxime','Ceftriaxone','Cefotaxime','Ceftazidime','Cefepime','CARBAPENEMS','Imipenem','Meropenem','Ertapenem','Doripenem','Colistin'],
                         AC.CONST_ORG_PSEUDOMONAS_AERUGINOSA:['Ceftazidime', 'Ciprofloxacin','Piperacillin/tazobactam','AMINOGLYCOSIDES','Gentamicin','Amikacin','CARBAPENEMS','Imipenem','Meropenem','Doripenem','Colistin'],
                         AC.CONST_ORG_ACINETOBACTER_BAUMANNII:['Tigecycline', 'Minocycline','AMINOGLYCOSIDES','Gentamicin','Amikacin','CARBAPENEMS','Imipenem','Meropenem','Doripenem','Colistin']}
CONST_V2_LIST_RENAME_ATB={"Oxacillin by MIC":"Oxacillin"}
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
def fn_rename_beforecombine(logger,df,scol,dict_rename={}):
    temp_df = df.copy(deep=True)
    for svalb in dict_rename:
        temp_df.loc[temp_df[scol]==svalb , scol] = dict_rename[svalb]
    return temp_df

def fn_combine_orgdata(logger,df,dict_combineorg={},lst_groupbycol=[],lst_sumcol=[],col_org="",col_totaln="",col_n="",col_percent="",col_lower95CI="",col_upper95CI="",sec4_5_totaln=0,bIsSec6=False):
    temp_df = df.copy(deep=True)
    #Sum dict
    dict_sum ={}
    for skey in lst_sumcol:
        dict_sum[skey] = ['sum']   
    #loop per combine org
    for scombine_org in dict_combineorg:
        lst_combine_org = dict_combineorg[scombine_org]
        #Filter only desire row
        df_filter = temp_df[temp_df[col_org].isin(lst_combine_org)]    
        #Add in combine org column as base for group by
        df_filter["combine_org"] = scombine_org
        df_filter = df_filter.groupby(by=lst_groupbycol + ["combine_org"],as_index=False).agg(dict_sum)
        df_filter.columns = list(map(''.join,df_filter.columns.values))
        #put data into dict
        for i in range(len(df_filter)):
            #Keep data for append back to df
            dict_data = {}
            for skey in temp_df.columns.tolist():
                dict_data[skey] = ""
            dict_data[col_org] = scombine_org
            for scol in lst_groupbycol:
                dict_data[scol] = df_filter.loc[i, scol]
            for scol in lst_sumcol:
                dict_data[scol] = df_filter.loc[i, scol + "sum"]    
            #if col_totaln is blank do not cal percent and CI
            if col_totaln != "":
                dict_data[col_percent] = "NA"
                dict_data[col_lower95CI] = "NA"
                dict_data[col_upper95CI] = "NA"
                try:
                    if df_filter.loc[i,col_totaln + "sum"]  != 0 :
                        if col_percent != "":
                            dict_data[col_percent] = round(df_filter.loc[i,col_n + "sum"] / df_filter.loc[i,col_totaln + "sum"] , 2) * 100
                        if col_lower95CI != "":
                            dict_data[col_lower95CI] = fn_wilson_lowerCI(x=df_filter.loc[i,col_n + "sum"], n=df_filter.loc[i,col_totaln + "sum"], conflevel=0.95, decimalplace=1)
                        if col_upper95CI != "":
                            dict_data[col_upper95CI] = fn_wilson_upperCI(x=df_filter.loc[i,col_n + "sum"], n=df_filter.loc[i,col_totaln + "sum"], conflevel=0.95, decimalplace=1)
                        if bIsSec6 == True:
                            dict_data[col_percent] = str(round(dict_data[col_percent],0)) + "% " + "(" + str(df_filter.loc[i,col_n + "sum"]) + "/" + str(df_filter.loc[i,col_totaln + "sum"]) + ")"
                except Exception as e:
                    AL.printlog("Error combine organism for compatible with V2.0 cal percent/CI for section 2,3: " +  str(e),True,logger) 
                    logger.exception(e)
            #if sec4_5_totaln > 0 will using col_n and sec4_5_totaln to cal percent and put in col_percent and cal CI put in the those 2 CI col.
            if int(sec4_5_totaln) > 0:
                dict_data[col_percent] = 0
                dict_data[col_lower95CI] = 0
                dict_data[col_upper95CI] = 0
                try:
                    if col_percent != "":
                        dict_data[col_percent] = (df_filter.loc[i,col_n + "sum"] / sec4_5_totaln)*AC.CONST_PERPOP
                    if col_lower95CI != "":
                        dict_data[col_lower95CI] = (AC.CONST_PERPOP/100)*fn_wilson_lowerCI(x=df_filter.loc[i,col_n + "sum"], n=sec4_5_totaln, conflevel=0.95, decimalplace=10)
                    if col_upper95CI != "":
                        dict_data[col_upper95CI] = (AC.CONST_PERPOP/100)*fn_wilson_upperCI(x=df_filter.loc[i,col_n + "sum"], n=sec4_5_totaln, conflevel=0.95, decimalplace=10)
                except Exception as e:
                    AL.printlog("Error combine organism for compatible with V2.0 cal percent/CI for section 2,3: " +  str(e),True,logger) 
                    logger.exception(e)
            df_append = pd.DataFrame([dict_data])
            temp_df =pd.concat([temp_df,df_append])
    return temp_df
def fn_filterorg(temp_df,col_org,lst_org):
    temp_df =temp_df[temp_df[col_org].isin(lst_org)]  
    temp_df[col_org] = pd.Categorical(temp_df[col_org], categories=lst_org, ordered=True)
    temp_df = temp_df.sort_values(col_org)
    return temp_df 
def fn_filterorg_atb(df,col_org,col_atb,dict_orgatb):
    lst_col = df.columns.tolist()
    result_df = pd.DataFrame(columns=lst_col)
    for sorg in dict_orgatb:
        lst_atb = dict_orgatb[sorg]
        temp_df = df[(df[col_org] == sorg) & (df[col_atb].isin(lst_atb))]
        temp_df[col_atb] = pd.Categorical(temp_df[col_atb], categories=lst_atb, ordered=True)
        temp_df = temp_df.sort_values(col_atb)
        result_df = pd.concat([result_df,temp_df],ignore_index = True)
        #result_df = pd.concat(temp_df)
    return result_df 

    
    
    
    
    

