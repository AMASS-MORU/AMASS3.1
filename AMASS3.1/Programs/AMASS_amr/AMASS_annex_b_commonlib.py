#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to analyze their own data and generate Supplementary data indicators reports systematically.

#Adding organisms to family_enterobacteriaceae at retrieve_all_family() (04/10/22)
#Editing prepare_datai_org() for supporting "; include" (04/10/22)
#Editing check_blocon() for supporting "; include" (04/10/22)
#Editing assign_org_popcon() for supporting "; include" (04/10/22)
#Editting create_summary_table_supp_poerr_v2() and create_summary_table_supp_pocon_v2() for supporting no case in df_mi (17/10/22)
#Editing prepare_datai_org() from "; include" to "; including"  (19/10/22)
#Editing check_blocon() from "; include" to "; including" (19/10/22)
#Editing assign_org_popcon() from "; include" to "; including" (19/10/22)

# Created on 20th April 2022
import psutil,os
import pandas as pd #for creating and manipulating dataframe
from pathlib import Path #for retrieving input's path
import AMASS_amr_const as AC
import AMASS_amr_commonlib as AL

def sub_printprocmem(sstate,logger) :
    try:
        process = psutil.Process(os.getpid())
        AL.printlog("Memory usage at state " +sstate + " is " + str(process.memory_info().rss) + " bytes.",False,logger) 
    except:
        AL.printlog("Error get process memory usage at " + sstate,True,logger)

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

def prepare_org_core(str_org, text_line=1, text_style="full", text_work="table", text_work_drug="N", text_bold="N"):
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
        if "spp" in lst_org[i] or "spp." in lst_org[i]:
            pass
        else:
            if text_work == "table":
                if text_work_drug == "N":  #S. pneumoniae (No durg information)
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

#Retrieving list of user's values
#return value: list of user's values
def retrieve_userlist(df_dict, amass_name, col_amass="amass_name", col_user="user_name"):
    lst = df_dict.loc[df_dict[col_amass]==amass_name,:].reset_index().loc[:,col_user].tolist()
    return [i for i in lst if i != ""]

#Merging antibiotic names and antibiotic class
#drug_class: df of list_of_antibiotics.xlsx
#drug_user: df of dictionary_for_microbiology_data (only antibiotics part)
#col_merge: column name for merging
#return: df of merged antibiotic name and antibiotic class
def merge_drug_drugclass(drug_class, drug_user, col_merge):
    return pd.merge(drug_class,drug_user,on=col_merge,how="outer").fillna("")

#Creating new column for AMASS' antibiotic names (xxx_clean_1)
#df: df of drug
#col_drug: column of original AMASS antibiotic names
#col_drug_new: column of new AMASS antibiotic names which will be created
#return: df of drug with additional column of new AMASS antibiotic names (xxx_clean_1)
def create_new_drugcol(df, col_drug, col_drug_new):
    df[col_drug_new] = ""
    for idx in df.index:
        if df.loc[idx,col_drug] == "":
            pass
        else:
            num = 1
            purpose_name = df.loc[idx,col_drug] + "_" + str(num)
            if purpose_name+"_clean" in set(df[col_drug_new]): #if purpose name is already used >>> add new purpose name
                num += 1
            else:
                pass
            df.loc[idx,col_drug_new] = df.loc[idx,col_drug] + "_" + str(num) + "_clean"
    return df

#Preparing organisms in datai
#df: datai
#col_org: column name of organism
#return value: dataframe with columns for "rule_organism", "except_organism", and "tax_level"
def prepare_datai_org(df, col_org, lst_fam, lst_ge, lst_sci): # + include
    df["organism_1"] = df[col_org].replace(regex=["organism_","family_"],value="").replace(regex=["_"],value=" ") 
    df[['rule_organism','except_organism']] = df['organism_1'].str.split(";",n=1,expand=True) #spliting org col to rule_organism and except_organism
    df['except_organism'] = df['except_organism'].str.replace(', ',',').str.replace(' except ','').str.replace('except ','').str.replace('except','').fillna('') #cleaning org spcies
    df[['rule_organism','include_organism']] = df['organism_1'].str.split(";",n=1,expand=True) #spliting org col to rule_organism and except_organism
    df['include_organism'] = df['include_organism'].str.replace(', ',',').str.replace(' including ','').str.replace('including ','').str.replace('including','').fillna('') #cleaning org spcies
    df["tax_level"] = ""
    for idx in df.index:
        ##Assigning tax_level
        if df.loc[idx,"rule_organism"].lower() == "all":
            df.at[idx,"tax_level"] = "all"
        elif df.loc[idx,"rule_organism"].replace(" spp","").lower() in lst_ge:
            df.at[idx,"tax_level"] = "organism"
        elif df.loc[idx,"rule_organism"].lower() in lst_fam:
            df.at[idx,"tax_level"] = "family"
        elif (df.loc[idx,"rule_organism"].lower() in lst_sci) or ("viridans" in df.loc[idx,"rule_organism"]):
            df.at[idx,"tax_level"] = "organism"
        else:
            df.at[idx,"tax_level"] = "NA"
        
        ##Cleaning except_organism and include_organism columns
        if "including" in df.loc[idx,"except_organism"]:
            df.at[idx,"except_organism"] = ""
        elif "except" in df.loc[idx,"include_organism"]:
            df.at[idx,"include_organism"] = ""
        else:
            pass
    return df

#worked!!
# def prepare_datai_org(df, col_org, lst_fam, lst_ge, lst_sci):
#     df["organism_1"] = df[col_org].replace(regex=["organism_","family_"],value="").replace(regex=["_"],value=" ")
#     df[['rule_organism','except_organism']] = df['organism_1'].str.split(";",1,expand=True) #spliting org col to rule_organism and except_organism
#     df['except_organism'] = df['except_organism'].str.replace(', ',',').str.replace(' except ','').str.replace('except ','').str.replace('except','').fillna('') #cleaning org spcies
#     df["tax_level"] = ""
#     for idx in df.index:
#         ##Assigning tax_level
#         if df.loc[idx,"rule_organism"].lower() == "all":
#             df.at[idx,"tax_level"] = "all"
#         elif df.loc[idx,"rule_organism"].replace(" spp","").lower() in lst_ge:
#             df.at[idx,"tax_level"] = "organism"
#         elif df.loc[idx,"rule_organism"].lower() in lst_fam:
#             df.at[idx,"tax_level"] = "family"
#         elif df.loc[idx,"rule_organism"].lower() in lst_sci:
#             df.at[idx,"tax_level"] = "organism"
#         else:
#             df.at[idx,"tax_level"] = "NA"
#     return df

#Preparing antibiotics in datai
#df_datai: datai
#df_drug: df_drug
#col_drug: column name of drug
#lst_drugclass: lst_cl
#lst_drugname: lst_dr
#return value: dataframe with columns for "except_antibiotic", and "amass_drug_rename" 
def prepare_datai_drug(df_datai, df_drug, col_drug, lst_drugclass, lst_drugname):
    df_datai[['antibiotic','except_antibiotic']] = df_datai[col_drug].str.split(";",n=1,expand=True) #spliting org col to rule_organism and except_organism
    df_datai['except_antibiotic'] = df_datai['except_antibiotic'].str.replace(', ',',').str.replace(' except ','').str.replace('except ','').str.replace('except','').fillna('') #cleaning org spcies
    df_datai["amass_drug_rename"] = ""
    for idx in df_datai.index:
        ##Assigning user_drug
        lst_anti = [x.replace(" ","") for x in df_datai.loc[idx,col_drug].split(",")]
        temp_1 = []
        temp_except = []
        temp_2 = ""
        if len(lst_anti) == 1:
            if lst_anti[0].lower() == "methicillin":
                temp_1 = df_drug.loc[df_drug["amass_drug"]=="Methicillin","amass_drug_rename"].tolist() +\
                        df_drug.loc[df_drug["amass_drug"]=="Oxacillin","amass_drug_rename"].tolist() + \
                        df_drug.loc[df_drug["amass_drug"]=="Cefoxitin","amass_drug_rename"].tolist()
            elif lst_anti[0] in lst_drugclass:
                temp_1 = df_drug.loc[df_drug["amass_class"]==lst_anti[0],"amass_drug_rename"].tolist() #list of full drug
                lst_except = [x.capitalize() if x!="penicillin_G" else "Penicillin_G" for x in df_datai.loc[idx,"except_antibiotic"].split(",")] #[] or ["Ceftazidime","Ceftriaxone"]
                temp_except = df_drug.loc[df_drug["amass_drug"].isin(lst_except),"amass_drug_rename"].tolist() #list of except drug
                temp_1 = list(set(temp_1)-set(temp_except)) #list of full - except
            elif lst_anti[0].lower() == "cephems":
                temp_1 = df_drug.loc[df_drug["amass_drug"]=="Cefazolin","amass_drug_rename"].tolist() +\
                        df_drug.loc[df_drug["amass_drug"]=="Cefuroxime","amass_drug_rename"].tolist() + \
                        df_drug.loc[df_drug["amass_drug"]=="Cefotaxime","amass_drug_rename"].tolist() + \
                        df_drug.loc[df_drug["amass_drug"]=="Ceftazidime","amass_drug_rename"].tolist() +\
                        df_drug.loc[df_drug["amass_drug"]=="Ceftriaxone","amass_drug_rename"].tolist() + \
                        df_drug.loc[df_drug["amass_drug"]=="Cefepime","amass_drug_rename"].tolist()
            elif lst_anti[0] in lst_drugname:
                temp_1 = df_drug.loc[df_drug["amass_drug"]==lst_anti[0],"amass_drug_rename"].tolist()
            else:
                pass
        elif len(lst_anti) > 1:
            temp_1_1 = []
            for i in range(len(lst_anti)):
                if lst_anti[i] in lst_drugclass:
                    temp_1_1 = df_drug.loc[df_drug["amass_class"]==lst_anti[i],"amass_drug_rename"].tolist() #list of full drug
                    lst_except = [x.capitalize() if x!="penicillin_G" else "Penicillin_G" for x in df_datai.loc[idx,"except_antibiotic"].split(",")] #[] or ["Ceftazidime","Ceftriaxone"]
                    temp_except = df_drug.loc[df_drug["amass_drug"].isin(lst_except),"amass_drug_rename"].tolist() #list of except drug
                    temp_1_1 = list(set(temp_1_1)-set(temp_except)) #list of full - except
                elif lst_anti[i] in lst_drugname:
                    temp_1_1 = df_drug.loc[df_drug["amass_drug"]==lst_anti[i],"amass_drug_rename"].tolist()
                temp_1 += temp_1_1
        else:
            pass
        temp_2 = ";".join([x for x in temp_1 if x != ""])
        df_datai.at[idx,"amass_drug_rename"] = temp_2
    return df_datai

#checking each record which has at least 1 AST result (for indicators 2, 3a, and 3b)
#df:               df of microbiology data
#col_ast:          column name for representing the results of checking AST
#str_ast_found:    string for assigning to col_ast when that record has at least 1 AST result
#str_ast_notfound: string for assigning to col_ast when that record has not any AST result
#return value:     df with col_ast
def check_ast_records(df, col_ast, str_ast_found, str_ast_notfound):
    df[col_ast] = ""
    for idx in df.index:
        temp_lst = df.loc[idx,:].tolist()
        if ('S' in temp_lst) or ("I" in temp_lst) or ("R" in temp_lst):
            df.at[idx,col_ast] = str_ast_found
        else:
            df.at[idx,col_ast] = str_ast_notfound
    return df

#Exporting Annex csv (Essential)
#df:            df of microbiology data
#df_blo_posneg: df of microbiology data (only blood and positive+negative cultures)
#df_blo_ast:    df of microbiology data (only blood and at least 1 AST result)
#rule1:         df of rule1
#rule2:         df of rule2
#rule3a:        df of rule3a
#rule3b:        df of rule3b
def export_annexB(df, df_blo_posneg, nogrowth_status, df_blo_ast, rule1, rule2, rule3a, rule3b):
    #Setting NA to annex_blood
    for col in df.columns:
        for idx in df.index:
            #Assigning "xx% (xx/xxx)" to all priority
            if idx == "blood_contamination":
                
                if nogrowth_status:
                    df.at[idx,"perc_"+col] = cal_perc_annex_v1(df.loc[idx,col],len(df_blo_posneg))
                    df.at[idx,"summary_"+col] = str(df.at[idx,"perc_"+col]) + "(" + str(int(df.loc[idx,col])) + "/" + str(len(df_blo_posneg)) + ")"
                else:
                    df.loc[idx,:] = df.loc[idx,:].astype(str)
                    df.at[idx,"Total"]    = "NA"
                    df.at[idx,"Critical"] = "NA"
                    df.at[idx,"High"]     = "NA"
                    df.at[idx,"Medium"]   = "NA"
                    df.at[idx,"perc_"+col] = "NA"
                    df.at[idx,"summary_"+col] = "NA"
            else:
                
                df.at[idx,"perc_"+col] = cal_perc_annex_v1(df.loc[idx,col],len(df_blo_ast))
                df.at[idx,"summary_"+col] = str(df.at[idx,"perc_"+col]) + " (" + str(int(df.loc[idx,col])) + "/" + str(len(df_blo_ast)) + ")"
            #Assigning "NA" to which priority that is not assigned to each indicator 1, 2, and 3 (in list_of_indicators.xlsx)
            if idx == "blood_contamination":
                lst_appeared_priority = list(set(rule1["priority"]))
            elif idx == "antibiotic_pathogen_combinations":
                lst_appeared_priority = list(set(rule2["priority"]))
            elif idx == "potential_errors":
                lst_appeared_priority = list(set(rule3a["priority"].tolist()+rule3b["priority"].tolist()))
            lst_notappeared_priority = assign_na_topriority(lst_appeared_priority)
            try:
                if len(lst_notappeared_priority) > 0:
                    for snsp in lst_notappeared_priority:
                        df.at[idx,snsp] = "NA"
            except Exception as e:
                print("Warning : " + str(e))
                pass
            """
            if len(lst_notappeared_priority) > 0:
                df.at[idx,lst_notappeared_priority] = "NA"
            else:
                pass
            """
    #Adding 2 columns for Total_blood_positive_AST(N), and Total_blood_positive_and_negative(N)
    df["Total_blood_positive_AST(N)"] = len(df_blo_ast)
    df["Total_blood_positive_and_negative(N)"] = len(df_blo_posneg)
    df = df.reset_index().rename(columns={"index":"Indicators", 
                                          "Total":"Total(N)", 
                                          "Critical":"Critical_priority(N)", 
                                          "High":"High_priority(N)", 
                                          "Medium":"Medium_priority(N)", 
                                          "summary_Total":"Total(%)",
                                          "summary_Critical":"Critical_priority(%)",
                                          "summary_High":"High_priority(%)", 
                                          "summary_Medium":"Medium_priority(%)"}).replace("NA (0/0)","NA")
    df.loc[:,"Indicators"] = ["blood_culture_contamination_rate","proportion_of_notifiable_antibiotic-pathogen_combinations","proportion_of_potential_errors_in_the_AST_results"]
    return df

#Creating df of annex by month
#df:            df of microbiology data
#df_blo_posneg: df of microbiology data (only blood and positive+negative cultures)
#df_blo_ast:    df of microbiology data (only blood and at least 1 AST result)
#col_spcdate:   column name of specimen_collection_date
#return value: dataframe of annex by month (frequency value)
def create_assign_annexB_bymonth(df, df_blo_posneg, df_blo_ast, col_spcdate):
    d_month = {1:"January", 2:"February", 3:"March", 4:"April", 5:"May", 6:"June", 7:"July", 8:"August", 9:"September", 10:"October", 11:"November", 12:"December", 
            "01":"January", "02":"February", "03":"March", "04":"April", "05":"May", "06":"June", "07":"July", "08":"August", "09":"September", "10":"October",  "11":"November", "12":"December", 
            "Jan":"January",   "Feb":"February",  "Mar":"March",    "Apr":"April", "May":"May",       "Jun":"June",      "Jul":"July",     "Aug":"August", "Sep":"September", "Oct":"October",   "Nov":"November", "Dec":"December", 
            "jan":"January",   "feb":"February",  "mar":"March",    "apr":"April", "may":"May",       "jun":"June",      "jul":"July",     "aug":"August", "sep":"September", "oct":"October",   "nov":"November", "dec":"December"}
    ##Based on denominator : positive blood culture
    #Change in V3.0 to use the date convert with common lib, fn_cleandate function 
    """
    #Grouping total micro_blood_dedup by month
    df["specimen_collection_date_fmt"] = pd.to_datetime(df[col_spcdate])
    df["month"] = df["specimen_collection_date_fmt"].dt.month
    df["count"] = 1
    ##Based on denominator : positive + negative blood culture
    #Grouping total micro_blood_dedup by month
    df_blo_posneg["specimen_collection_date_fmt"] = pd.to_datetime(df_blo_posneg[col_spcdate])
    df_blo_posneg["month"] = df_blo_posneg["specimen_collection_date_fmt"].dt.month
    df_blo_posneg["count"] = 1
    #Grouping total micro_blood_dedup_ast by month
    df_blo_ast["specimen_collection_date_fmt"] = pd.to_datetime(df_blo_ast[col_spcdate])
    df_blo_ast["month"] = df_blo_ast["specimen_collection_date_fmt"].dt.month
    df_blo_ast["count"] = 1
    """
    #Grouping total micro_blood_dedup by month
    df["month"] = df[col_spcdate].dt.month
    df["count"] = 1
    ##Based on denominator : positive + negative blood culture
    #Grouping total micro_blood_dedup by month
    df_blo_posneg["month"] = df_blo_posneg[col_spcdate].dt.month
    df_blo_posneg["count"] = 1
    #Grouping total micro_blood_dedup_ast by month
    df_blo_ast["month"] = df_blo_ast[col_spcdate].dt.month
    df_blo_ast["count"] = 1
    #----------------------------------------------------------------------------
    #Counting number of warning of each month
    df_blo_posneg_bymonth = df_blo_posneg.loc[:,["month","count"]].groupby("month").count().rename(index=d_month,columns={"count":"total"}) #Based on denominator : positive + negative blood culture
    df_blo_ast_bymonth    = df_blo_ast.loc[:,["month","count"]].groupby("month").count().rename(index=d_month,columns={"count":"total_ast"})
    df_blo_bymonth_indi_1 = df.loc[df["warning_indicator_1"] != "",["month","count"]].groupby("month").count().rename(index=d_month,columns={"count":"indicator_1"})
    df_blo_bymonth_indi_2 = df.loc[df["warning_indicator_2"] != "",["month","count"]].groupby("month").count().rename(index=d_month,columns={"count":"indicator_2"})
    df_blo_bymonth_indi_3 = df.loc[(df["warning_indicator_3a"] != "") | (df["warning_indicator_3b"] != ""),["month","count"]].groupby("month").count().rename(index=d_month,columns={"count":"indicator_3"})
    df_blo_bymonth = pd.DataFrame(index=["January","February","March","April","May","June","July","August","September","October","November","December"])
    return pd.concat([df_blo_bymonth, 
                      df_blo_bymonth_indi_1, 
                      df_blo_bymonth_indi_2, 
                      df_blo_bymonth_indi_3, 
                      df_blo_posneg_bymonth, 
                      df_blo_ast_bymonth],axis=1).fillna(0).astype(int)

#Exporting annexB by month
#df: df of microbiology data
#return value: df of annexB by month
def export_annexB_bymonth(df, nogrowth_status, indicator_1="indicator_1", indicator_2="indicator_2", indicator_3="indicator_3"):
    #Calculating perc_indicator_1, perc_indicator_2, and perc_indicator_3
    for idx in df.index:
        for col in [indicator_1,indicator_2,indicator_3]:
            col_total = ""
            #Calculating %
            if col == indicator_1: #assigning total for indicator_1
                col_total = "total"
                if nogrowth_status:
                    df.at[idx,"perc_"+col]    = cal_perc_annex_v1(df.loc[idx,col],df.loc[idx,col_total])
                    df.at[idx,"summary_"+col] = str(df.loc[idx,"perc_"+col]) + " (" + str(df.loc[idx,col]) + "/" + str(df.loc[idx,col_total]) + ")"
                    df.at[idx,"summary_"+col] = df.loc[idx,"summary_"+col].replace("NA (0/0)","NA")
                else:
                    df.at[idx,:]            =  df.loc[idx,:].astype(str)
                    df.at[idx,col]            = "NA"
                    df.at[idx,"perc_"+col]    = "NA"
                    df.at[idx,"summary_"+col] = "NA"

            else:
                col_total = "total_ast"
                df.at[idx,"perc_"+col] = cal_perc_annex_v1(df.loc[idx,col],df.loc[idx,col_total])
                df.at[idx,"summary_"+col] = str(df.loc[idx,"perc_"+col]) + " (" + str(df.loc[idx,col]) + "/" + str(df.loc[idx,col_total]) + ")"
                df.at[idx,"summary_"+col] = df.loc[idx,"summary_"+col].replace("NA (0/0)","NA")
    return df.reset_index().rename(columns={"index":"month", 
                                            "total":"Total_blood_positive_negative(N)", 
                                            "total_ast":"Total_blood_positive_AST(N)", 
                                            indicator_1:"blood_culture_contamination_rate(N)", 
                                            indicator_2:"proportion_of_notifiable_antibiotic-pathogen_combinations(N)",
                                            indicator_3:"proportion_of_potential_errors_in_the_AST_results(N)",
                                            "perc_"+indicator_1:"blood_culture_contamination_rate(%)", 
                                            "perc_"+indicator_2:"proportion_of_notifiable_antibiotic-pathogen_combinations(%)",
                                            "perc_"+indicator_3:"proportion_of_potential_errors_in_the_AST_results(%)",
                                            "summary_"+indicator_1:"summary_blood_culture_contamination_rate(%)", 
                                            "summary_"+indicator_2:"summary_proportion_of_notifiable_antibiotic-pathogen_combinations(%)",
                                            "summary_"+indicator_3:"summary_proportion_of_potential_errors_in_the_AST_results(%)"})

#Exporting records with warning
#df: dataframe of blood records with warning
#dict_datai: dictionary for linking indicator names
#str_filename_withstatus: filename of exported df (with report status)
#str_filename_withoutstatus: filename of exported df (without report status)
#col_spcnum: column name of specimen_number
#return value: no return value; exported excel file of microbiology data with warning 
def export_records_withwarning_wide(df, 
                                    dict_datai, 
                                    str_filename_withstatus, 
                                    str_filename_withoutstatus, 
                                    col_hn, 
                                    col_spcdate, 
                                    col_spcnum):
    drop_col = [x for x in list(df.columns) if "_clean" in x]
    df = df.drop(columns=drop_col+["mapped_blood","mapped_culture","mapped_fam","priority_indicator_1","priority_indicator_2","priority_indicator_3a","priority_indicator_3b","combine"])
    if (pd.isna(col_spcnum)) or (col_spcnum == "") or (col_spcnum not in df.columns):
        df = df.drop(columns=["mapped_specimen_number"])
    else:
        pass
    df = df.rename(columns={"warning_indicator_1":"1",
                            "warning_indicator_2":"2",
                            "warning_indicator_3a":"3a",
                            "warning_indicator_3b":"3b"}).sort_values(by=[col_hn, "1", "2", "3a", "3b"]).rename(columns=dict_datai)
    #Change in V3.0 to use the date convert with common lib, fn_cleandate function 
    df_withstatus = format_date_forexportation(df = df,col_date = col_spcdate)
    # df_withstatus.to_excel(str_filename_withstatus,index=False)
    df_withstatus.to_csv(str_filename_withstatus,index=False)
    df_withoutstatus = df_withstatus.drop(columns=["mapped_spctype","mapped_sci","mapped_gen","status_indicator_1","status_indicator_2","status_indicator_3a","status_indicator_3b","car_3gc","flu_3gc"])
    df_withoutstatus["Notifiable antibiotic-pathogen combination"] = df_withoutstatus["Notifiable antibiotic-pathogen combination"].replace(regex=["markednewline\("],value=" (")
    df_withoutstatus["Potential error in either identification or AST result: the species identified usually exhibits intrinsic resistant to the antibiotic but AST result suggested susceptible"] = df_withoutstatus["Potential error in either identification or AST result: the species identified usually exhibits intrinsic resistant to the antibiotic but AST result suggested susceptible"].replace(regex=["markednewline\("],value=" (")
    df_withoutstatus["Discordant AST results"] = df_withoutstatus["Discordant AST results"].replace(regex=["markednewline\("],value=" (")
    # df_withoutstatus.to_excel(str_filename_withoutstatus,index=False)
    df_withoutstatus.to_csv(str_filename_withoutstatus,index=False)

#Exporting records with warning (long format)
#df: dataframe of blood records with warning
#dict_datai: dictionary for linking indicator names
#str_filename_withstatus: filename of exported df (with report status)
#str_filename_withoutstatus: filename of exported df (without report status)
#col_hn: column name of hn
#col_spcdate: column name of specimen_collection_date
#col_spctype: column name of specimen type
#col_organism: column name of organism
#col_spcnum: column name of specimen_number
#return value: no return value; exported excel file of microbiology data with warning 
def export_records_withwarning_long(df, 
                                    dict_datai, 
                                    str_filename_rawmicro,
                                    str_filename_withstatus, 
                                    str_filename_withoutstatus,
                                    col_hn,
                                    col_spcdate,
                                    col_spctype,
                                    col_organism,
                                    col_spcnum):
    list_micro_warning = df["combine"].tolist()
    drop_col = [x for x in list(df.columns) if "_clean" in x]
    micro_onlywarning_1 = df.reset_index().drop(columns=drop_col+["priority_indicator_1","priority_indicator_2","priority_indicator_3a","priority_indicator_3b"])
    micro_raw = pd.read_excel(str_filename_rawmicro).fillna("")
    if (pd.isna(col_spcnum)) or (col_spcnum == "") or (col_spcnum not in df.columns): #If there is no available specimen_nember column >>> 
        micro_raw           = create_columncombine(df=micro_raw,          hn=col_hn,date=col_spcdate,type_=col_spctype,org=col_organism)
        micro_onlywarning_1 = create_columncombine(df=micro_onlywarning_1,hn=col_hn,date=col_spcdate,type_=col_spctype,org=col_organism)
    else:
        micro_raw = create_columncombine(df=micro_raw,hn=col_hn,date=col_spcdate,type_=col_spctype,num=col_spcnum,org=col_organism)
    micro_onlywarning_1 = micro_onlywarning_1.set_index("combine")

    micro_raw[["warning_indicator_1","status_indicator_1","warning_indicator_2","status_indicator_2",
        "warning_indicator_3a","status_indicator_3a","warning_indicator_3b","status_indicator_3b"]] = ""
    micro_onlywarning_2 = pd.DataFrame()
    micro_onlywarning_3 = pd.DataFrame()
    for idx_warn in micro_onlywarning_1.index:
        y = micro_raw.loc[micro_raw["combine"]==idx_warn]
        temp_df = pd.DataFrame(columns=micro_raw.columns)
        temp_idx = 0
        for idx_raw in y.index:
            if len(micro_onlywarning_1.loc[idx_warn,])==len(micro_onlywarning_1.columns): #If there is 1 record : 1 warning >>> direct assign
                for col in ["warning_indicator_1","status_indicator_1","warning_indicator_2","status_indicator_2","warning_indicator_3a","status_indicator_3a","warning_indicator_3b","status_indicator_3b"]:
                    y.at[idx_raw,col] = micro_onlywarning_1.loc[idx_warn,col]
            else:
                for col in ["warning_indicator_1","status_indicator_1","warning_indicator_2","status_indicator_2","warning_indicator_3a","status_indicator_3a","warning_indicator_3b","status_indicator_3b"]:
                    warn = micro_onlywarning_1.loc[idx_warn,col].tolist()
                    for w in warn:
                        if y.loc[idx_raw,col] != "":
                            temp_df.at[temp_idx,:] = y.loc[idx_raw,:]
                            temp_df.at[temp_idx,col] = w
                            temp_idx += 1
                        else:
                            y.at[idx_raw,col] = w
                    y = pd.concat([y,temp_df])
        micro_onlywarning_2 = pd.concat([micro_onlywarning_2,y]).reset_index().drop(columns=["index"])
        micro_onlywarning_3 = df.copy().rename(columns={"warning_indicator_1":"1",
                                                        "warning_indicator_2":"2",
                                                        "warning_indicator_3a":"3a",
                                                        "warning_indicator_3b":"3b"}).sort_values(by=[col_hn, "1", "2", "3a", "3b"]).rename(columns=dict_datai).drop(columns=["priority_indicator_1",
                                                        "priority_indicator_2",
                                                        "priority_indicator_3a",
                                                        "priority_indicator_3b",
                                                        "combine"])
    micro_onlywarning_3= format_date_forexportation(df = micro_onlywarning_3,col_date = col_spcdate)
    # micro_onlywarning_3.to_excel(str_filename_withstatus,index=False)
    micro_onlywarning_3.to_csv(str_filename_withstatus,index=False)
    micro_onlywarning_2 = micro_onlywarning_2.rename(columns={"warning_indicator_1":"1",
                                                            "warning_indicator_2":"2",
                                                            "warning_indicator_3a":"3a",
                                                            "warning_indicator_3b":"3b"}).sort_values(by=[col_hn, "1", "2", "3a", "3b"]).rename(columns=dict_datai).drop(columns=[
                                                            "status_indicator_1",
                                                            "status_indicator_2",
                                                            "status_indicator_3a",
                                                            "status_indicator_3b", 
                                                            "combine"])
    micro_onlywarning_2["Notifiable antibiotic-pathogen combination"] = micro_onlywarning_2["Notifiable antibiotic-pathogen combination"].replace(regex=["markednewline\("],value=" (").replace(regex=["markedstartitalic","markedenditalic"],value="")
    micro_onlywarning_2["Potential error in either identification or AST result: the species identified usually exhibits intrinsic resistant to the antibiotic but AST result suggested susceptible"] = micro_onlywarning_2["Potential error in either identification or AST result: the species identified usually exhibits intrinsic resistant to the antibiotic but AST result suggested susceptible"].replace(regex=["markednewline\("],value=" (").replace(regex=["markedstartitalic","markedenditalic"],value="")
    micro_onlywarning_2["Discordant AST results"] = micro_onlywarning_2["Discordant AST results"].replace(regex=["markednewline\("],value=" (").replace(regex=["markedstartitalic","markedenditalic"],value="")
    micro_onlywarning_2 = format_date_forexportation(df = micro_onlywarning_2,col_date = col_spcdate)
    # micro_onlywarning_2.to_excel(str_filename_withoutstatus,index=False)
    micro_onlywarning_2.to_csv(str_filename_withoutstatus,index=False)

#Exporting records of annex A
#return value: df of annexA
def export_records_annexA(df, dictionary, col_spcdate, col_organism):
    #Retrieving records for notifiable organisms
    drop_col = [x for x in list(df.columns) if "_clean" in x]
    dict_org_A = dictionary.loc[dictionary["amass_name"].isin(["organism_burkholderia_pseudomallei", 
                                                                "organism_corynebacterium_diphtheriae", 
                                                                "organism_neisseria_gonorrhoeae",
                                                                "organism_neisseria_meningitidis",
                                                                "organism_streptococcus_suis"])|
                                dictionary["amass_name"].str.contains("organism_brucella_")|
                                dictionary["amass_name"].str.contains("organism_salmonella_")|
                                dictionary["amass_name"].str.contains("organism_vibrio_")].reset_index().drop(columns=['index']).fillna("")

    lst_dict_org = dict_org_A.loc[dict_org_A["user_name"]!="","user_name"].tolist()
    df_A = df.loc[df[col_organism].isin(lst_dict_org)]
    df_A["mapped_spctype"] = df_A["mapped_spctype"].replace({"blood":"Blood","csf":"CSF","genital swab":"Genital swab","rts":"RTS","stool":"Stool","urine":"Urine","others":"Others"})
    df_A = format_date_forexportation(df = df_A,col_date = col_spcdate)
    df_A.drop(columns=drop_col).to_excel(AC.CONST_PATH_REPORTWITH_PID +"Report_with_patient_identifiers_annexA_withstatus.xlsx",header=True,index=False)
    if "mapped_specimen_number" in df_A.columns: #drop columns before exportation
        df_A = df_A.drop(columns=["mapped_specimen_number"])
    else:
        pass
    df_A = df_A.drop(columns=drop_col+["mapped_spctype","mapped_blood","mapped_culture","mapped_fam","mapped_sci","mapped_gen"])
    df_A.to_excel(AC.CONST_PATH_REPORTWITH_PID +"Report_with_patient_identifiers_annexA.xlsx",header=True,index=False)


#Retrieving available family based on dictionary_for_microbiology_data
#return: dataframe with 2 columns; amass_name: family name, user_name: scientific name
def retrieve_all_family(df_dict):
    list_org_ent = ["organism_citrobacter_amalonaticus"    , "organism_citrobacter_freundii"     , "organism_citrobacter_koseri",    "organism_citrobacter_spp",
                "organism_enterobacter_cloacae_complex", "organism_enterobacter_spp"         ,
                "organism_escherichia_coli"            , "organism_escherichia_hermannii"    , "organism_escherichia_spp", 
                "organism_hafnia_alvei"                ,
                "organism_klebsiella_aerogenes"        , "organism_klebsiella_oxytoca"       , "organism_klebsiella_pneumoniae", "organism_klebsiella_pseudopneumoniae" , "organism_klebsiella_variicola", "organism_klebsiella_spp",
                "organism_leclercia_adecarboxylata"    , 
                "organism_morganella_morganii"         , "organism_morganella_spp"           ,
                "organism_proteus_mirabilis"           , "organism_proteus_penneri"          , "organism_proteus_rettgeri",      "organism_proteus_stuartii"            , "organism_proteus_vulgaris"    , "organism_proteus_spp",
                "organism_providencia_spp"             ,
                "organism_raoultella_ornitholytica"    , "organism_raoultella_planticola"    , "organism_raoultella_terrigena",  "organism_raoultella_spp", 
                "organism_salmonella_agona", 
                "organism_salmonella_amsterdam"        ,  "organism_salmonella_anatum"       , "organism_salmonella_arechavaleta",
                "organism_salmonella_bareilly"         , "organism_salmonella_blockley"      , "organism_salmonella_bongori", 
                "organism_salmonella_bonnA"             , "organism_salmonella_bovismorbificans", "organism_salmonella_braenderup", 
                "organism_salmonella_brandenburg"      , "organism_salmonella_bredeney"      , "organism_salmonella_cerro", 
                "organism_salmonella_chester"          , "organism_salmonella_choleraesuis"  , "organism_salmonella_copenhagen", 
                "organism_salmonella_derby"            , "organism_salmonella_dublin"        , "organism_salmonella_emek", 
                "organism_salmonella_enteritidis"      , "organism_salmonella_falkensee"     , "organism_salmonella_gallinarum", 
                "organism_salmonella_give"             , "organism_salmonella_goldcoast"     , "organism_salmonella_hadar", 
                "organism_salmonella_heidelberg"       , "organism_salmonella_hirschfeldii"  , "organism_salmonella_infantis",
                "organism_salmonella_java"             , "organism_salmonella_javiana"       , "organism_salmonella_kaapstad", 
                "organism_salmonella_kedougou"         , "organism_salmonella_kentucky"      , "organism_salmonella_kottbus", 
                "organism_salmonella_krefeld"          , "organism_salmonella_lexington"     , "organism_salmonella_litchfield", 
                "organism_salmonella_livingstone"      , "organism_salmonella_lomita"        , "organism_salmonella_london",
                "organism_salmonella_manhattan"        , "organism_salmonella_mbandaka"      , "organism_salmonella_montevideo", 
                "organism_salmonella_muenchen"         , "organism_salmonella_muenster"      , "organism_salmonella_narashino", 
                "organism_salmonella_newport"          , "organism_salmonella_ohio"          , "organism_salmonella_oranienburg", 
                "organism_salmonella_orion"            , "organism_salmonella_ouakam"        , "organism_salmonella_panama", 
                "organism_salmonella_paratyphi"        , "organism_salmonella_poona"         , "organism_salmonella_potsdam", 
                "organism_salmonella_pullorum"         , "organism_salmonella_reading"       , "organism_salmonella_rissen", 
                "organism_salmonella_saintpaul"        , "organism_salmonella_schottmuelleri", "organism_salmonella_schwarzengrund", 
                "organism_salmonella_senftenberg"      , "organism_salmonella_stanley"       , "organism_salmonella_tennessee", 
                "organism_salmonella_thompson"         , "organism_salmonella_typhi"         , "organism_salmonella_typhimurium", 
                "organism_salmonella_typhisuis"        , "organism_salmonella_virchow"       , "organism_salmonella_virginia", 
                "organism_salmonella_wandsworth"       , "organism_salmonella_weltevreden"   , "organism_salmonella_weybridge", 
                "organism_salmonella_worthington"      , "organism_non-typhoidal_salmonella_spp", "organism_salmonella_spp", 
                "organism_serratia_marcescens"         , "organism_serratia_spp"             , 
                "organism_shigella_boydii"             , "organism_shigella_dysenteriae"     , "organism_shigella_flexneri"      , "organism_shigella_sonnei"             , "organism_shigella_spp",      
                "organism_yersinia_enterocolitica"     , "organism_yersinia_pseudotuberculosis" , "organism_yersinia_spp",
                "organism_other_enterobacteriaceae"]
    ##Retrieving family Enterobacteriaceae
    dict_org_ent = df_dict.loc[df_dict["amass_name"].isin(list_org_ent),["amass_name","user_name"]].fillna("")
    dict_org_ent["amass_name"] = "family_enterobacteriaceae"
    dict_org_ent = dict_org_ent.loc[dict_org_ent["user_name"]!=""]
    ##Retrieving other family
    dict_org_oth = df_dict.loc[(df_dict["amass_name"].str.contains("organism_other"))&(df_dict["user_name"]!=""),["amass_name","user_name"]].fillna("")
    dict_org_oth = dict_org_oth.loc[dict_org_oth["user_name"]!=""]
    for index in dict_org_oth.index:
        dict_org_oth.loc[index,"amass_name"] = dict_org_oth.loc[index,"amass_name"].replace("organism_other_","family_")
    return pd.concat([dict_org_ent,dict_org_oth],axis=0)

#Retrieving available scientific name based on dictionary_for_microbiology_data
#return: dataframe with 2 columns; amass_name: scientific name which are contained "organism_" as prefix, user_name: scientific name from user
def retrieve_ava_scientific_name(df_dict):
    dict_org_sci = df_dict.loc[df_dict["amass_name"].str.contains("organism_"),["amass_name","user_name"]].fillna("") #dataframe containing organism name
    return dict_org_sci.loc[dict_org_sci["user_name"]!=""].reset_index().drop(columns=["index"])

#Selecting value from list based on rule
#return: list containing values without "removed_value"
def select_value(list_raw, removed_value=""):
    return [x for x in list_raw if x != removed_value]

#return value: list of intersec priority
def assign_na_topriority(lst_pri):
    priority =  ["Critical","High","Medium"]
    lst_diff = list(set(priority).difference(lst_pri))
    for i in range(len(lst_diff)):
        lst_diff[i] = "summary_" + lst_diff[i]
    return(lst_diff) #setA - setB

def check_assign_priority(priority_micro,priority_qc):
    priority_temp = [priority_micro,priority_qc]
    priority_verify = ""
    if "Critical" in priority_temp:
        priority_verify = "Critical"
    elif "Critical" not in priority_temp:
        if "High" in priority_temp:
            priority_verify = "High"
        elif "High" not in priority_temp:
            if "Medium" in priority_temp:
                priority_verify = "Medium"
            else:
                priority_verify = ""
    return priority_verify

def check_assign_reportstatus(status_micro,status_qc):
    status_temp = [status_micro,status_qc]
    status_verify = ""
    if "yes" in status_temp:
        status_verify = "yes"
    elif "yes" not in status_temp:
        if "no" in status_temp:
            status_verify = "no"
        else:
            status_verify = ""
    return status_verify

#Retrieving user value from dictionary_for_microbiology_data.xlsx
#dict_df is df of dictionary_for_microbiology_data.xlsx.
#col_amass is column name of amass value of dictionary_for_microbiology_data.xlsx.
#col_user is column name of user value of dictionary_for_microbiology_data.xlsx.
#amass_name is amass value that will be used to retrieve user value.
#return value is string of user value.
def retrieve_uservalue(dict_df, amass_name, col_amass="amass_name", col_user="user_name"):
    return dict_df.loc[dict_df[col_amass]==amass_name,:].reset_index().loc[0,col_user]

##Selecting available drug for rule2,3, and 5
def select_ava_drug(lst_qc_drug,lst_micro_col):
    lst = []
    for drug in lst_qc_drug:
        if drug in lst_micro_col:
            lst.append(drug)
        else:
            pass
    return lst

#search by key to map value
def create_dict_for_map(lst_key,value):
    d_ = {}
    for key in lst_key:
        d_[key] = value
    return d_

#Mapping AST result (In the case that microbiology_data is contained a mixture of ast formats)
#df_micro: dataframe of microbiology_data
#df_drug: dataframe of mapped drug_list
#d_ast: dictionary for mapping ast result (key is user's ast, value is amass' ast)
#return value: dataframe of microbiology_data with containing amass ast format
def map_ast_result(df_micro,df_drug,d_ast):
    d_colname = {}
    for idx in df_drug.index:
        num = 1
        if df_drug.loc[idx,"user_drug"] == "":
            pass
        else:
            if df_drug.loc[idx,"user_drug"] in df_micro.columns: #in case: appear in dictionary but not appear in microbiology_data
                mapped_col = df_drug.loc[idx,"amass_drug_rename"]
                if mapped_col+"_"+str(num) in df_micro.columns:
                    num += 1
                else:
                    pass
                df_micro[mapped_col+"_"+str(num)] = df_micro[df_drug.loc[idx,"user_drug"]].map(d_ast).fillna("").astype("category") #mapping ast result
                # df_micro[mapped_col+"_"+str(num)] = df_micro[df_drug.loc[idx,"user_drug"]].map(d_ast).fillna("") #mapping ast result
                d_colname[mapped_col+"_"+str(num)] = mapped_col
    df_micro = df_micro.rename(columns=d_colname) #renaming col_name
    return df_micro

##Mapping AMASS' family name, scientific name and genus name
#return value: dataframe of microbiology_data with mapped AMASS' family, scientific, and genus name
def map_fam_org_gen_name(df_micro, col_org, df_fam, df_sci):
    df_micro["mapped_fam"] = ""
    df_micro["mapped_sci"] = ""
    df_micro["mapped_gen"] = ""
    # df_micro["mapped_gen"] = ""
    for idx_mi in df_micro.index:
        fam = ""
        sci = ""
        ##Family name
        if df_micro.loc[idx_mi,col_org] in df_fam["user_name"].tolist():
            fam = list(df_fam.loc[df_fam["user_name"]==df_micro.loc[idx_mi,col_org],"amass_name"])[0]
            df_micro.at[idx_mi,"mapped_fam"] = fam.replace("family_","")
        else:
            pass
        ##Organism name
        if df_micro.loc[idx_mi,col_org] in df_sci["user_name"].tolist():
            org = list(df_sci.loc[df_sci["user_name"]==df_micro.loc[idx_mi,col_org],"amass_name"])[0]
            df_micro.at[idx_mi,"mapped_sci"] = org.replace("organism_","").replace("_"," ")
        else:
            pass
        ##Genus name
        if df_micro.loc[idx_mi,col_org] in df_sci["user_name"].tolist():
            gen = list(df_sci.loc[df_sci["user_name"]==df_micro.loc[idx_mi,col_org],"amass_name"])[0]
            df_micro.at[idx_mi,"mapped_gen"] = gen.split("_")[1] + " spp"
        else:
            pass
    return df_micro

#Creating list for drugs using in discordant in rule2
#Return value: list of drugs
def create_list_for_rule2_comb(lst_drug_ava):
    lst_drug_ava_1 = []
    if len(lst_drug_ava) == 0:
        pass
    else:
        lst_drug_ava_1 = [i for i in lst_drug_ava[0]]
    return lst_drug_ava_1

#Assigning rule and except organisms for potential contaminants
#df_qc: raw dataframe of potential contaminants
#return value: dataframe with rule and except organisms
def assign_org_pocon(df_qc):
    df_qc[["rule_ge","rule_sp"]] = ""
    df_qc["except_sp"] = ""
    df_qc["include_sp"] = ""
    for idx in df_qc.index:
        rule_org = df_qc.loc[idx,"rule_organism"].split(" ")
        except_org = df_qc.loc[idx,"except_organism"].split(",")
        include_org = df_qc.loc[idx,"include_organism"].split(",")
        if df_qc.loc[idx,"tax_level"] == "organism":
            df_qc.loc[idx,["rule_ge"]] = rule_org[0] #string
            df_qc.loc[idx,["rule_sp"]] = rule_org[1] #string
            df_qc.at[idx,"except_sp"]  = except_org  #list
            df_qc.at[idx,"include_sp"] = include_org #list
        elif df_qc.loc[idx,"tax_level"] == "family":
            df_qc.loc[idx,["rule_ge"]] = rule_org[0] #string
            df_qc.loc[idx,["rule_sp"]] = "spp"       #string
            df_qc.at[idx,"except_sp"]  = except_org  #list
            df_qc.at[idx,"include_sp"] = include_org #list
    return df_qc

#worked!!
# def assign_org_pocon(df_qc):
#     df_qc[["rule_ge","rule_sp"]] = ""
#     df_qc["except_sp"] = ""
#     for idx in df_qc.index:
#         rule_org = df_qc.loc[idx,"rule_organism"].split(" ")
#         except_org = df_qc.loc[idx,"except_organism"].split(",")
#         if df_qc.loc[idx,"tax_level"] == "organism":
#             df_qc.loc[idx,["rule_ge"]] = rule_org[0] #string
#             df_qc.loc[idx,["rule_sp"]] = rule_org[1] #string
#             df_qc.at[idx,"except_sp"] = except_org   #list
#         elif df_qc.loc[idx,"tax_level"] == "family":
#             df_qc.loc[idx,["rule_ge"]] = rule_org[0] #string
#             df_qc.loc[idx,["rule_sp"]] = "spp"       #string
#             df_qc.at[idx,"except_sp"] = except_org   #list
#     return df_qc

#Assigning available drugs for other sets
#df_qc: raw dataframe of other sets
#return value: dataframe with available drugs
def assign_drug_oth(df_micro, df_qc):
    df_qc["amass_drug_ava"] = ""
    for idx_qc in df_qc.index:
        drugname = df_qc.loc[idx_qc,"amass_drug_rename"].split(";")
        micro_col = df_micro.columns.tolist()
        lst_drugname = select_ava_drug(drugname,micro_col) #Selecting available drug for rule2
        if lst_drugname == []:
            df_qc.at[idx_qc,"amass_drug_ava"] = ""
        else:
            df_qc.at[idx_qc,"amass_drug_ava"] = lst_drugname
    df_qc = df_qc.loc[df_qc["amass_drug_ava"] != ""]
    return df_qc

#Retrieving index from micro_pos dataframe that are under rule_specimen (compatible with rule1)
#df_micro: dataframe of microbiology_data
#df_rule: dataframe of rule
#return value: index of a dataframe which are NEEDED to process
def retrieve_idx_basedon_spc(df_micro, df_rule):
    lst_spc = list(set(df_rule["antibiotic"].replace(regex="_specimen",value="")))
    return df_micro.loc[df_micro["mapped_spctype"].isin(lst_spc),:].index

#Retrieving records based on specimen types which are set in list_of_indicators.xlsx
#Return value: A dataframe of records that specimen types matched to list_of_indicators.xlsx 
def retrieve_record_basedon_spc(df, lst_spc):
    if "all_specimen" in lst_spc:
        df = df
    else:
        lst_spc_1 = [x.replace("_specimen","") for x in lst_spc] #["blood_specimen"] >>> ["blood"]
        df = df.loc[df["mapped_spctype"].isin(lst_spc_1)]
    return df

#Retrieving index from micro_pos dataframe that are rule_organism (compatible with rule2, rul3a, and rule3b)
#df_micro: dataframe of microbiology_data
#df_rule: dataframe of rule
#return value: index of a dataframe which are NEEDED to process
def retrieve_idx_basedon_org(df_micro, df_rule):
    df_micro_1 = pd.DataFrame()
    lst_fam = []
    lst_sci = []
    if "all" in list(set(df_rule["tax_level"])):
        df_micro_1 = df_micro.copy()
    else:
        lst_fam = list(set(df_rule.loc[df_rule["tax_level"]=="family","rule_organism"]))
        lst_sci = list(set(df_rule.loc[df_rule["tax_level"]=="organism","rule_organism"]))
        df_micro_1 = df_micro.loc[(df_micro["mapped_fam"].isin(lst_fam))|(df_micro["mapped_sci"].isin(lst_sci)),:]
    return df_micro_1.index

#Selecting dataframe of indicator by organism from microbiology_data
#return value: selected dataframe of indicator based on microbiology_data
def select_indicator_by_org_v2(df_micro, idx_mi, df_qc, col_qc_org="rule_organism"):
    lst_org = [org for org in df_micro.loc[idx_mi,["mapped_fam","mapped_sci","mapped_gen"]].tolist() if org !=""]
    # genus = df_micro.loc[idx_mi,"mapped_sci"].split(" ")[0] + " spp"
    # lst_org = lst_org + [genus]
    if "all" in df_qc[col_qc_org].tolist(): #if there is "all" in lst_org >>> add "all" to lst_org
        lst_org = ["all"] + lst_org
    else:
        pass
    return df_qc.loc[df_qc[col_qc_org].isin(lst_org),:]

#Checking organism for rule2, 3a, and 3b
#tax_qc: tax_level from datai
#fam_mi: family name from microbiology_data
#fam_qc: family name from datai
#org_mi: organism name from microbiology_data
#org_qc: organism_name from datai
def check_org(tax_qc,org_qc,fam_mi,org_mi):
    warning  = ""
    if tax_qc.lower() == "all": # ALL
        warning = True
    elif tax_qc.lower() == "family":
        if fam_mi.lower() == org_qc.lower():
            warning = True
        else:
            warning = False
    elif tax_qc.lower() == "organism":  
        if org_mi == org_qc:
            warning = True
        else:
            if " spp" in org_qc:
                org_mi = org_mi.split(" ")
                org_qc = org_qc.split(" ")
                if org_mi[0] == org_qc[0]:
                    warning = True
            else:
                warning = False        
    else:
        warning  = False
    return warning

#Checking blood culture contaminant
#return value: boolean (True: assign warning, False: passed)
def check_blocon(qc_ge,  qc_ex, qc_in, qc_spc, mi_sci, mi_fam, mi_spc):
    boolean_result = False
    ##Species## Rule: staphylococcus spp; except staphylococcus aureus
    if (qc_ge.lower() in mi_sci.lower()) or (qc_ge.lower() in mi_fam.lower()): #If "staphylococcus" in "staphylococcus spp" >>> do next process
        if mi_sci.lower() in qc_ex:                                        #If "staphylococcus aureus" in ["staphylococcus aureus", "staphylococcus ludunensis"] >>> passed
            pass
        else:                                                              #If "staphylococcus spp" NOT in ["staphylococcus aureus", "staphylococcus ludunensis"] >>> do next process
            if check_spctype_1(qc_spc,mi_spc):                             #if result_spc is True >>> assign priority and report status
                boolean_result = True
            else:                                                          #if result_spc is FALSE >>> pass
                pass

    if (mi_sci != "") and (mi_sci.lower() in qc_in):           #If "streptococcus gordonii" in "streptococcus gordonii" >>> do next process
        if check_spctype_1(qc_spc,mi_spc):                             #if result_spc is True >>> assign priority and report status
                boolean_result = True
        else:                                                          #if result_spc is FALSE >>> pass
            pass
    else:
        pass
    return boolean_result

#worked!!    
# def check_blocon(qc_ge,  qc_ex,  qc_spc, mi_sci, mi_fam, mi_spc):
#     boolean_result = False
#     ##Species## Rule: staphylococcus spp; except staphylococcus aureus
#     if qc_ge.lower() in mi_sci.lower() or qc_ge.lower() in mi_fam.lower(): #If "staphylococcus" in "staphylococcus spp" >>> do next process
#         if mi_sci.lower() in qc_ex:                                        #If "staphylococcus aureus" in ["staphylococcus aureus", "staphylococcus ludunensis"] >>> passed
#             pass
#         else:                                                              #If "staphylococcus spp" NOT in ["staphylococcus aureus", "staphylococcus ludunensis"] >>> do next process
#             if check_spctype_1(qc_spc,mi_spc):                             #if result_spc is True >>> assign priority and report status
#                 boolean_result = True
#             else:                                                          #if result_spc is FALSE >>> pass
#                 pass
#     else:
#         pass
#     return boolean_result

#Checking potential error in AST result and intrinsically resistant to antibiotic
#return value: boolean (True: assign warning, False: passed)
def check_poerr_pathogen(tax_qc, org_qc, fam_mi, org_mi, lst_amr, criteria):
    boolean_warning = ""
    if len(lst_amr) > 0 and list(set(lst_amr)) != [""]: #If drug is available >>> do next process
        if criteria == "NS":                            #If criteria is "NS" : lst_amr = ["","R"] or ["R","S"] or ["R","R"] or ["","I"] or ["I","S"] or ["I","R"] or ["I","I"] >>> do next process
            if "R" in lst_amr or "I" in lst_amr:        #If lst_amr = ["R","S"] or ["I","S"] >>> do next process
                boolean_warning = check_org(tax_qc, org_qc, fam_mi, org_mi)
            else:                                       #If lst_amr = ["","S"] >>> passed
                pass
        elif criteria == "S":                           #If criteria is "S" : lst_amr = ["","S"] or ["R","S"] or ["I","S"] or ["S","S"] >>> do next process
            if "R" in lst_amr or "I" in lst_amr:        #If lst_amr = ["R","S"] or ["I","S"] >>> pass
                pass
            else:                                       #If lst_amr = ["","S"] or ["S","S"] >>> do next process
                boolean_warning = check_org(tax_qc, org_qc, fam_mi, org_mi)
        else:
            pass
    else:                                               #lst_amr = [] >>> passed
        pass
    return boolean_warning

#Checking discordant AST result
#return value: boolean (True: assign warning, False: passed)
def check_disast(tax_qc, org_qc, fam_mi, org_mi, lst_antibiotic):
    boolean_warning = False
    if "S" in lst_antibiotic and ("R" in lst_antibiotic or "I" in lst_antibiotic): #lst_amr = ["S","R"] or ["S","I"] >>> do next process
        boolean_warning = check_org(tax_qc, org_qc, fam_mi, org_mi)
    else:                                               #lst_amr = ["","R"] or ["","S"] or ["","I"] >>> passed
        pass
    return boolean_warning

#Checking discordant AST result by fixing antibiotic-S and antibiotic-NS
#return value: boolean (True: assign warning, False: passed)
def check_disast_fixS(lst_amr_S, lst_amr_NS):
    boolean_warning = False
    if "R" in lst_amr_NS or "I" in lst_amr_NS:
        if "S" in lst_amr_S and ("R" not in lst_amr_S and "I" not in lst_amr_S):
            boolean_warning = True
        else:
            pass
    else:
        pass
    return boolean_warning

def print_round(count, total):
    if count % 1000 == 0:
        print ("Processed samples: " + str(count) + "/" + str(total) + "(" + str(round(count/total*100,2)) + " %" + ")" )
    else:
        pass

#Checking whether specimen type is set for indicator1 or not
#spc_qc: specimen type from dataqc
#spc_mi: specimen type from microbiology_data
#return value: True (is set for consideration) or False (is NOT set for consideration)
def check_spctype_1(spc_qc,spc_mi):
    result = ""
    if spc_qc.lower() == "all":
        result = True
    elif spc_qc.lower() == spc_mi.lower():
        result = True
    else:
        result = False
    return result

#Creating new column name "combine" that combining 5 columns (hn,spcdate,spctype,spcnum,organism) together
#df is the original df
#hn is hospital number.
#date is specimen date.
#type_ is specimen type.
#num is specimen number.
#organism is organism name
#return value is the original df with combine column
def create_columncombine(df,hn,date,type_,org,num=""):
    if num != "":
        df["combine"] = df[hn].astype(str)+";"+df[date].astype(str)+";"+df[type_].astype(str)+";"+df[num].astype(str)+";"+df[org].astype(str)
    else:
        df["combine"] = df[hn].astype(str)+";"+df[date].astype(str)+";"+df[type_].astype(str)+";"+df[org].astype(str)
    return df

#Retrieving records from the original df based on warning from AT LEAST 1 rule
#df is the original df for retrieving
#col_rule1, col_rule2, col_rule3a, col_rule3b are warning column name of rule1-3
#return value is the original df with INCLUDING at least 1 warning rule
def retrieve_recordbywarning(df):
    return df.loc[(df["warning_indicator_1"]!="")|(df["warning_indicator_2"]!="")|(df["warning_indicator_3a"]!="")|(df["warning_indicator_3b"]!="")]

#Creating table for Annex B on page41
#Return value: table for Annex B on page41
def create_assign_annex_v2(df):
    rule = ["indicator_1", "indicator_2", "indicator_3"]
    priority = ["Critical", "High", "Medium"]
    annex = pd.DataFrame(index=rule, columns=["Total"]+priority).fillna(0)
    for r in rule:
        if r == "indicator_3":
            for pri in priority:
                temp_df = df.loc[(df["warning_"+r+"a"]!="")|(df["warning_"+r+"b"]!="")]
                annex.at[r,pri] = len(df.loc[(df["priority_"+r+"a"]==pri)|(df["priority_"+r+"b"]==pri)])
        else:
            for pri in priority:
                annex.at[r,pri] = len(df.loc[(df["warning_"+r]!="") & (df["priority_"+r]==pri)])
    annex["Total"] = annex[priority[0]] + annex[priority[1]] + annex[priority[2]]
    return annex

#Calculating percentage for Annex B
#Return value: percentage in correct format
def cal_perc_annex_v1(int_numerator,int_denominate):
    try:
        temp_perc = round(int(int_numerator)/int(int_denominate)*100,1)
        if temp_perc < 0.05: #numerator is always >= 0
            temp_result = "0%"
        elif temp_perc >= 0.05 and temp_perc < 0.1:
            temp_result = "<0.1%"
        elif temp_perc >= 0.1 and temp_perc < 0.95:
            temp_result = str(round(temp_perc,1)) + "%"
        else:
            temp_result = str(round(temp_perc)) + "%"
    except:
        temp_result = "NA"
    return temp_result

#Creating summary table for Supplementary report (for potential error inAST results)
#return value: dataframe of each set of indicator with %blood samples
def create_summary_table_supp_poerr_v2(df_qc, df_mi, int_denominator, col_mi_drug, col_qc_org="rule_organism", col_qc_drug="antibiotic", col_mi_gen="mapped_gen", col_mi_org="mapped_sci", col_mi_fam="mapped_fam"):
    df_qc_1 = df_qc.copy()
    df_qc_1[["number_blood","blood_samples"]] = ""
    for idx in df_qc_1.index:
        rule_org = df_qc_1.loc[idx,col_qc_org]
        rule_drug = df_qc_1.loc[idx,col_qc_drug]
        if rule_org.lower() == "all":
            df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[df_mi[col_mi_drug]==rule_drug])
        else:
            if len(df_mi)>0:
                df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[((df_mi[col_mi_org]==rule_org)&(df_mi[col_mi_drug]==rule_drug))|((df_mi[col_mi_fam]==rule_org)&(df_mi[col_mi_drug]==rule_drug))|((df_mi[col_mi_gen]==rule_org)&(df_mi[col_mi_drug]==rule_drug))])
            else:
                df_qc_1.at[idx,"number_blood"] = 0
        df_qc_1.at[idx,col_qc_org] = rule_org.capitalize().replace(" spp"," spp.")
        df_qc_1.at[idx,col_qc_drug] = df_qc_1.loc[idx,col_qc_drug].replace("_"," ")
        df_qc_1.at[idx,"blood_samples"] = cal_perc_annex_v1(df_qc_1.loc[idx,"number_blood"],int_denominator) + \
                                                    " (" + str(df_qc_1.loc[idx,"number_blood"]) + "/" + str(int_denominator) + ")"
    return df_qc_1.loc[:,[col_qc_org,col_qc_drug,"blood_samples","number_blood"]].reset_index().drop(columns=["index","number_blood"])


#old
# def create_summary_table_supp_poerr_v2(df_qc, df_mi, int_denominator, col_mi_drug, col_qc_org="rule_organism", col_qc_drug="antibiotic", col_mi_gen="mapped_gen", col_mi_org="mapped_sci", col_mi_fam="mapped_fam"):
#     df_qc_1 = df_qc.copy()
#     df_qc_1[["number_blood","blood_samples"]] = ""
#     for idx in df_qc_1.index:
#         rule_org = df_qc_1.loc[idx,col_qc_org]
#         rule_drug = df_qc_1.loc[idx,col_qc_drug]
#         if rule_org.lower() == "all":
#             df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[df_mi[col_mi_drug]==rule_drug])
#         else:
#             df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[((df_mi[col_mi_org]==rule_org)&(df_mi[col_mi_drug]==rule_drug))|((df_mi[col_mi_fam]==rule_org)&(df_mi[col_mi_drug]==rule_drug))|((df_mi[col_mi_gen]==rule_org)&(df_mi[col_mi_drug]==rule_drug))])
#         df_qc_1.at[idx,col_qc_org] = rule_org.capitalize().replace(" spp"," spp.")
#         df_qc_1.at[idx,col_qc_drug] = df_qc_1.loc[idx,col_qc_drug].replace("_"," ")
#         df_qc_1.at[idx,"blood_samples"] = cal_perc_annex_v1(df_qc_1.loc[idx,"number_blood"],int_denominator) + \
#                                                     " (" + str(df_qc_1.loc[idx,"number_blood"]) + "/" + str(int_denominator) + ")"
#     return df_qc_1.loc[:,[col_qc_org,col_qc_drug,"blood_samples","number_blood"]].reset_index().drop(columns=["index","number_blood"])

def create_summary_table_supp_poerr(df_qc, df_mi, int_denominator, col_mi_drug, col_qc_org="rule_organism", col_qc_drug="antibiotic", col_mi_org="mapped_sci", col_mi_fam="mapped_fam"):
    df_qc_1 = df_qc.copy()
    df_qc_1[["number_blood","blood_samples"]] = ""
    for idx in df_qc_1.index:
        rule_org = df_qc_1.loc[idx,col_qc_org]
        rule_drug = df_qc_1.loc[idx,col_qc_drug]
        if rule_org.lower() == "all":
            df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[df_mi[col_mi_drug]==rule_drug])
        else:
            df_qc_1.at[idx,"number_blood"] = len(df_mi.loc[((df_mi[col_mi_org]==rule_org)&(df_mi[col_mi_drug]==rule_drug))|((df_mi[col_mi_fam]==rule_org)&(df_mi[col_mi_drug]==rule_drug))])
        df_qc_1.at[idx,col_qc_org] = rule_org.capitalize()
        df_qc_1.at[idx,col_qc_drug] = df_qc_1.loc[idx,col_qc_drug].replace("_"," ")
        df_qc_1.at[idx,"blood_samples"] = cal_perc_annex_v1(df_qc_1.loc[idx,"number_blood"],int_denominator) + \
                                                    " (" + str(df_qc_1.loc[idx,"number_blood"]) + "/" + str(int_denominator) + ")"
    return df_qc_1.loc[:,[col_qc_org,col_qc_drug,"blood_samples","number_blood"]].reset_index().drop(columns=["index","number_blood"])

#Creating summary table for Supplementary report (for potential blood contaminant)
#return value: dataframe of each set of indicator with %blood samples
def create_summary_table_supp_pocon(df_qc, df_mi, int_denominator, col_mi_warning="warning_indicator_1", col_qc_org="rule_organism", col_qc_exorg="except_organism", col_mi_org="mapped_sci", col_mi_fam="mapped_fam"):
    df_mi_1 = df_mi.copy().loc[df_mi[col_mi_warning]!=""]
    df_mi_1[["genus","species"]] = df_mi_1[col_mi_org].str.split(" ",n=1,expand=True).fillna("")
    lst_ge = list(set(df_mi_1["genus"])) + list(set(df_mi_1[col_mi_fam])) #list of unique genus + family

    df_qc_1 = df_qc.copy()
    df_qc_1[["genus","species"]] = df_qc_1.loc[:,col_qc_org].str.split(" ",n=1,expand=True)
    df_qc_1["number_blood"]=0
    df_qc_1["blood_samples"]=""
    for idx in df_qc_1.index:
        str_ge = df_qc_1.loc[idx,"genus"]
        if str_ge in lst_ge:
            df_qc_1.at[idx,"number_blood"] = len(df_mi_1.loc[df_mi_1["genus"]==str_ge])
        df_qc_1.at[idx,"blood_samples"] = cal_perc_annex_v1(df_qc_1.loc[idx,"number_blood"],int_denominator) + \
                                                        " (" + str(df_qc_1.loc[idx,"number_blood"]) + "/" + str(int_denominator) + ")"
    return df_qc_1.loc[:,[col_qc_org,col_qc_exorg,"blood_samples","number_blood"]].reset_index().drop(columns=["index","number_blood"])

def create_summary_table_supp_pocon_v2(boolean_nogrowth,df_qc, df_mi, int_denominator, col_qc_org="rule_organism", col_qc_exorg="except_organism", col_qc_inorg="include_organism", col_mi_org="mapped_sci", col_mi_fam="mapped_fam"):
    lst_ge = []
    df_mi_1 = df_mi.copy()
    if  len(df_mi) > 0: #there is at least 1 record
        df_mi_1[["genus","species"]] = df_mi_1[col_mi_org].str.split(" ",n=1,expand=True)
        lst_ge = list(set(df_mi_1["genus"])) + list(set(df_mi_1[col_mi_fam])) #list of unique genus + family
    else:
        lst_ge = []
    df_qc_1 = df_qc.copy()
    df_qc_1[col_qc_org] = df_qc_1[col_qc_org].replace(regex=[" spp"], value=" spp.")
    df_qc_1[["genus","species"]] = df_qc_1.loc[:,col_qc_org].str.split(" ",n=1,expand=True)
    df_qc_1["number_blood"]=0
    df_qc_1["blood_samples"]=""
    for idx in df_qc_1.index:
        str_ge = df_qc_1.loc[idx,"genus"]
        if boolean_nogrowth is True:
            if len(df_mi_1)>0:
                if (str_ge in lst_ge) and ("viridans" not in str_ge):
                    df_qc_1.at[idx,"number_blood"] = len(df_mi_1.loc[df_mi_1["genus"]==str_ge])
                elif "viridans" in str_ge:
                    df_qc_1.at[idx,"number_blood"] = len(df_mi_1.loc[df_mi_1["genus"].str.contains("viridans")]) + len(df_mi_1.loc[df_mi_1["genus"]=="streptococcus"])
                else:
                    pass
                df_qc_1.at[idx,"blood_samples"] = cal_perc_annex_v1(df_qc_1.loc[idx,"number_blood"],int_denominator) + \
                                                                " (" + str(df_qc_1.loc[idx,"number_blood"]) + "/" + str(int_denominator) + ")"
            else:
                df_qc_1.at[idx,"blood_samples"] = "0 (0/" + str(int_denominator) + ")"
        else:
            df_qc_1.at[idx,"blood_samples"] = "NA"
    return df_qc_1.loc[:,[col_qc_org,col_qc_exorg,col_qc_inorg, "blood_samples","number_blood"]].reset_index().drop(columns=["index","number_blood"])

#Formating date (i.e. specimen collection date) into "01 Jan 2012" format
#Return value: dataframe with formated date column
def format_date_forexportation(df,col_date):
    # df[col_date] = pd.to_datetime(df[col_date])
    for idx in df.index:
        if isinstance(df.loc[idx,col_date], str):
            pass
        else:
            df.at[idx,col_date]=df.loc[idx,col_date].strftime("%d %b %Y")
    return df