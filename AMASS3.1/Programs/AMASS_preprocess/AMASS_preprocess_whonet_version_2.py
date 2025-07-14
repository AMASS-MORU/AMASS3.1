#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable dictionary for microbiology data in the right format for running AMASS version 2.0 via WHONET systematically.

# Created on 20th April 2022
import logging #for creating error_log
import pandas as pd #for creating dataframe
from AMASS_preprocess_function_version_2 import * #for importing preprocess functions

# Create a logging instance
logger = logging.getLogger('AMASS_preprocess_whonet_version_2.py')
logger.setLevel(logging.INFO)
# Assign a file-handler to that instance
fh = logging.FileHandler("./error_preprocess_whonet.txt")
fh.setLevel(logging.ERROR)
# Format your logs (optional)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
# Add the handler to your logging instance
logger.addHandler(fh)

path = "./"
head_dict_2 = "Variable names used for \"antibiotics\" in AMASS"
head_dict_3 = "Data values of variable used for \"specimen_type\" in AMASS"

try:
    config = pd.read_excel(path + "Configuration/Configuration.xlsx")
    if check_config(config, "preprocess_function"):
        #Reading dictionary_for_microbiology_data.xlsx
        dict_micro = pd.DataFrame()
        try:
            try:
                dict_micro = pd.read_excel(path + "dictionary_for_microbiology_data.xlsx").fillna("")
            except:
                try:
                    dict_micro = pd.read_csv(path + "dictionary_for_microbiology_data.csv").fillna("")
                except:
                    dict_micro = pd.read_csv(path + "dictionary_for_microbiology_data.csv", encoding="windows-1252").fillna("")
            dict_micro.columns = ["amass_name","user_name","requirement","explanation"]
            #Retrieving section for antibiotics
            idx_micro_drug = dict_micro.iloc[:,0].tolist().index(head_dict_2) #index of part2 header
            idx_micro_spc  = dict_micro.iloc[:,0].tolist().index(head_dict_3) #index of part3 header
            df_drug = dict_micro.iloc[idx_micro_drug+1:idx_micro_spc].reset_index().drop(columns=["index"])
            #Checking which dictionary_for_microbiology_data.xlsx is needed to preprocess
            if len(df_drug.loc[(df_drug["user_name"].str.contains("_N"))|
                            (df_drug["user_name"].str.contains("_E"))|
                            (df_drug["user_name"].str.contains("_F"))]) > 0:
                #Reading microbiology_data.xlsx
                micro_raw = pd.DataFrame()
                try:
                    micro_raw = pd.read_excel(path + "microbiology_data.xlsx").fillna("")
                except:
                    try:
                        micro_raw = pd.read_csv(path + "microbiology_data.csv").fillna("")
                    except:
                        micro_raw = pd.read_csv(path + "microbiology_data.csv",encoding="windows-1252").fillna("")
                #Retrieving column names 
                df_col_micro = pd.DataFrame(micro_raw.columns.tolist(),columns=["micro_name"])
                #Merging df_drug and df_col_micro
                df_merge = pd.merge(df_drug.iloc[:,:2],df_col_micro,left_on="user_name",right_on="micro_name",how="inner")
                #Selecing available antibitoics which are presented in columns of microbiology_data.xlsx
                df_drug_sel = df_drug.loc[df_drug["user_name"].isin(df_merge["micro_name"])]
                #Creating and exporting new dicttionary_for_microbiology_data.xlsx
                df_info = dict_micro.iloc[:idx_micro_drug+1] #part 1
                df_oth = dict_micro.iloc[idx_micro_spc:] #part 3-8
                df_dict = pd.concat([df_info,df_drug_sel,df_oth]).reset_index().drop(columns=["index"])
                df_dict.to_excel(path + "dictionary_for_microbiology_data.xlsx",index=False,header=["\"Don't change content in this column, but you can add rows with the same content if needed\"",
                                                                                            "\"Change content in this column to represent how variable names, antibiotic names, or variable values are written in your raw microbiology data file\"", 
                                                                                            "Requirements", 
                                                                                            "Explanations"])
            else:
                pass
        except Exception as e:
            logger.exception(e) # Will send the errors to the file
            pass
    else:
        pass
except Exception as e:
    logger.exception(e) # Will send the errors to the file
    pass
