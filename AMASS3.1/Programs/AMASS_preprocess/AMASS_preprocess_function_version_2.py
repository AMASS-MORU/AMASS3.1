#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable microbiology data in the right format for running AMASS version 2.0 systematically.

# Created on 20th April 2022
import pandas as pd #for creating dataframe
import numpy as np #for checking dimentions of dataframe

def retrieve_uservalue(dict_df, amass_name, col_amass="amass_name", col_user="user_name"):
    return dict_df.loc[dict_df[col_amass]==amass_name,:].reset_index().loc[0,col_user]

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
    
#Checking file_format for checking that microbiology_data is needed for either reformatting or not
#return : boolean value; True if need reformatting; False if NO need reformatting
def check_need_to_reformat(dict_micro):
    lst = dict_micro.iloc[:,0].tolist()
    fmt = dict_micro.loc[lst.index("file_format"),"user_name"]
    result = ""
    if fmt == "long":
        result = True
    else:
        result = False
    return result

#Checking available value of list_1 that is NOT found in list_1
#return: list of non-matched value of list_1
def check_ava_column(list_1,list_2):
    return list(set(list_1)-set(list_2))

#Removing duplicated rows
#return: dataframe with deduplicated specimen information
def deduplicate(data_micro):
    return data_micro.drop_duplicates(subset='combine',keep='first')

#Creating dataframe 'drug_uni' with #row=len(micro_uni) and #column=len(drug_uni)
#return: blank dataframe in wide format
def create_blank_df_widefmt(data_micro,lst_drug):
    df_drug = pd.DataFrame(index=range(len(data_micro)),columns=list(lst_drug))
    return pd.concat([data_micro,df_drug],axis=1)

#Filling ast_result to blank dataframe
#return: dataframe with ast_result
def fill_ast_to_rawmicro(data_micro_amass, data_micro_raw, col_drug, col_ast):
    count_round = 0
    for idx_amass in data_micro_amass.index: #for loop based on index
        var_df = data_micro_raw.loc[data_micro_raw["combine"]==idx_amass,[col_drug,col_ast]] #selecting sub-dataframe based on index
        if np.shape(var_df) == (2,): #unique record --var_df is contained 2 values (2,)
            data_micro_amass.at[idx_amass,var_df[0]] = var_df[1] #assigning sensitivity 
        else: #multiple records --var_df is contained >=2 rows and 2 columns (x,2)
            for row in range(len(var_df)): #for loop based on range of rows
                data_micro_amass.at[idx_amass,var_df.iloc[row,0]] = var_df.iloc[row,1] #assigning sensitivity
        count_round += 1
        if count_round%1000 == 0:
            print ("Finish reformatting : " + str(count_round) + " lines")
    return data_micro_amass