#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable microbiology data in the right format for running AMASS version 2.0 systematically.

# Created on 20th April 2022
import logging #creating error_log
import pandas as pd #creating dataframe
from AMASS_preprocess_function_version_2 import * #for importing preprocess functions

# Create a logging instance
logger = logging.getLogger('AMASS_preprocess_version_2.py')
logger.setLevel(logging.INFO)
# Assign a file-handler to that instance
fh = logging.FileHandler("./error_preprocess.txt")
fh.setLevel(logging.ERROR)
# Format your logs (optional)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
# Add the handler to your logging instance
logger.addHandler(fh)

try:
    config = pd.read_excel("./Configuration/Configuration.xlsx")
    if check_config(config, "preprocess_function"):
        dict_raw = pd.DataFrame()
        try:
            try:
                dict_raw = pd.read_excel("./dictionary_for_microbiology_data.xlsx").iloc[:,:2].fillna("")
            except:
                try:
                    dict_raw = pd.read_csv("./dictionary_for_microbiology_data.csv").iloc[:,:2].fillna("")
                except:
                    dict_raw = pd.read_csv("./dictionary_for_microbiology_data.csv", encoding="windows-1252").iloc[:,:2].fillna("")
            dict_raw.columns = ["amass_name", "user_name"]

            if check_need_to_reformat(dict_raw): #long format
                dict_lst = dict_raw.iloc[:,0].tolist()
                hn      = retrieve_uservalue(dict_raw, "hospital_number")
                spcnum  = retrieve_uservalue(dict_raw, "specimen_number")
                spcdate = retrieve_uservalue(dict_raw, "specimen_collection_date")
                spctype = retrieve_uservalue(dict_raw, "specimen_type")
                organism= retrieve_uservalue(dict_raw, "organism")
                antibiotic = retrieve_uservalue(dict_raw, "antibiotic")
                astresult  = retrieve_uservalue(dict_raw, "ast_result")
                nogrowth   = retrieve_uservalue(dict_raw, "organism_no_growth")

                data_micro_raw = pd.DataFrame()
                try:
                    data_micro_raw = pd.read_excel("./microbiology_data.xlsx").fillna("")
                except:
                    try:
                        data_micro_raw = pd.read_csv("./microbiology_data.csv").fillna("")
                    except:
                        data_micro_raw = pd.read_csv("./microbiology_data.csv",encoding="windows-1252").fillna("")
                try: #required columns are founded in microbiology_data
                    data_micro_ast = data_micro_raw.loc[(data_micro_raw[hn]!="")&(data_micro_raw[spcdate]!="")&(data_micro_raw[spctype]!="")&(data_micro_raw[organism]!="")&(data_micro_raw[spcnum]!="")&(data_micro_raw[antibiotic]!="")&(data_micro_raw[astresult]!=""),:]
                    #Creating new column 'combine' that containing hn;spcdate;spctype;organism
                    data_micro_ast['combine'] = data_micro_ast[hn].astype(str)+';'+data_micro_ast[spcdate].astype(str)+';'+ \
                                        data_micro_ast[spctype].astype(str)+';'+data_micro_ast[organism].astype(str)+';'+ \
                                        data_micro_ast[spcnum].astype(str)
                    data_micro_dedup = deduplicate(data_micro_ast).reset_index().loc[:,[hn,spcdate,spctype,organism,spcnum,'combine']]
                    lst_ava_drug = pd.unique(data_micro_ast.loc[:,antibiotic])
                    data_micro_amass = create_blank_df_widefmt(data_micro_dedup,lst_ava_drug).set_index("combine")
                    data_micro_amass_1 = fill_ast_to_rawmicro(data_micro_amass, data_micro_ast, antibiotic, astresult).reset_index().drop(columns=["combine"]).fillna("")
                    print ("Writing microbiology_data.xlsx (wide format)")
                    data_micro_final=pd.concat([data_micro_amass_1,data_micro_raw.loc[data_micro_raw[astresult]==""]])
                    data_micro_final.to_excel('./microbiology_data_reformatted.xlsx',index=False) #Exporting microbiology_data
                except Exception as e:
                    logger.exception(e) # Will send the errors to the file
                    pass

            else:
                pass
        except Exception as e:
            logger.exception(e) # Will send the errors to the file
            pass
except Exception as e:
    logger.exception(e) # Will send the errors to the file
    pass
