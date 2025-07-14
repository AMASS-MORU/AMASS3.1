
#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** CONST FILE and Configurations                                                                   ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: CHALIDA RAMGSIWUTISAK
# Created on: 30 AUG 2023
# Last update on: 20 JUNE 2025 #v3.1 3102
import AMASS_amr_const as AC
import pandas as pd #for creating and manipulating dataframe
CONST_RUN_ANNEXC_WITH_NOHOSP = False
CONST_COL_PROFILE    = "profile"
CONST_COL_PROFILEID  = "profile_ID"
CONST_COL_PROFILETEMP= "profile_temp"
CONST_COL_DIFFSPCADM = "annexc_diff_spcdate_admdate"
CONST_COL_AMRPATHOGEN= "annexc_amrpathogen"
CONST_COL_ANNEXCCOHO = "annexc_InfOri"
CONST_NEWVARNAME_ASTVAN_RIS = AC.CONST_NEWVARNAME_PREFIX_RIS + "Vancomycin"
#columns for Configuration
CONST_COL_AMASS_PRMNAME= "AMASS basic run"
CONST_COL_USER_PRMNAME = "Setting parameters"
#columns for AMASS dictionaries
CONST_COL_AMASSNAME = "amass_name"
CONST_COL_USERNAME = "user_name"
#columns for SATSCAN input.case
CONST_COL_TESTGROUP = "test_group"
CONST_COL_SPCDATE = "specimen_collection_date"
CONST_COL_RESISTPROFILE = "resistant_profile"
CONST_COL_CASE = "num_case"
CONST_COL_WEEK = "num_week"
CONST_COL_WARDID = "ward_ID"
CONST_COL_WARDNAME = "ward_name"
#columns for SATSCAN results.col
CONST_COL_LOCID = "LOC_ID"
CONST_COL_SDATE = "START_DATE"
CONST_COL_EDATE = "END_DATE"
CONST_COL_TSTAT = "TEST_STAT"
CONST_COL_PVAL  = "P_VALUE"
CONST_COL_OBS   = "OBSERVED"
CONST_COL_EXP   = "EXPECTED"
CONST_COL_ODE   = "ODE"
CONST_COL_CLEANSDATE = "clean_sdate"
CONST_COL_CLEANEDATE = "clean_edate"
CONST_COL_CLEANPVAL  = "clean_pval"
CONST_COL_WARDPROFILE = "ward_profile_id"
#columns for SATSCAN AnnexC_listofclusters_XXXX_XXX.xlsx
CONST_COL_NEWSDATE = "start signal date"
CONST_COL_NEWEDATE = "end signal date"
CONST_COL_NEWPVAL  = "p-value"
CONST_COL_NEWOBS   = "observed cases"
#columns for SATSCAN AnnexC_counts_by_organism_XXX.xlsx
CONST_COL_NUMOBS  = "Number_of_observed_cases"
CONST_COL_NUMWARD = "Number_of_wards"
#values for AMASS microbiology_data.xlsx
CONST_PRENAME_PROFILEID = "profile_"
CONST_VALUE_WARD = "ward"
#values for ast_information.xlsx
CONST_VALUE_PASSEDATB   ="P" #v3.1 3102
CONST_VALUE_NOTPASSEDATB="F" #v3.1 3102
CONST_VALUE_NOTTESTEDATB="-" #v3.1 3102
CONST_VALUE_RESISTANT   ="R" #v3.1 3102
CONST_VALUE_INTERMEDIATE="I" #v3.1 3102
CONST_VALUE_SUSCEPTIBLE ="S" #v3.1 3102
CONST_VALUE_NA        ="NA" #v3.1 3102
CONST_VALUE_C1_1      ="C1:≥" #v3.1 3102
CONST_VALUE_C1_2      ="%" #v3.1 3102
CONST_VALUE_C2        ="C2:variation" #v3.1 3102
CONST_VALUE_SUMMARY   ="Summary" #v3.1 3102
lst_ris   = [CONST_VALUE_RESISTANT,CONST_VALUE_INTERMEDIATE,CONST_VALUE_SUSCEPTIBLE] #v3.1 3102
lst_ris_na= [CONST_VALUE_RESISTANT,CONST_VALUE_INTERMEDIATE,CONST_VALUE_SUSCEPTIBLE,CONST_VALUE_NA] #v3.1 3102
#columns for profile_information.xlsx
CONST_COL_NUMPROFILE_ALL="No. of patients with a clinical specimen culture positive"
CONST_COL_NUMPROFILE_BLO="No. of patients with blood culture positive"
#columns for SATSCAN Graphs
CONST_COL_SWEEKDAY = "startweekday"
CONST_COL_OTHWARD  = "Other wards"
#columns for Report1 result
CONST_COL_DATAFILE  = "Type_of_data_file"
CONST_COL_PARAM     = "Parameters"
CONST_COL_DATE      = "Values"
CONST_VALUE_DATAFILE= "overall_data"
CONST_VALUE_SDATE   = "Minimum_date"
CONST_VALUE_EDATE   = "Maximum_date"
#Supplementary report
CONST_NUM_SLICEROW_PROFILE = 20
CONST_NUM_SLICECOL_PROFILE = 13
CONST_NUM_INCH = 72
CONST_LST_COLWIDTH_PROFILE = [0.8*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,
                              0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,
                              0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH,0.7*CONST_NUM_INCH]
#values for profiling
# CONST_VALUE_TESTATBRATE = "tested_antibiotic_rate"
# CONST_VALUE_RRATE = "resistant_rate"
# CONST_VALUE_IRATE = "intermediate_rate"
# CONST_VALUE_SRATE = "susceptible_rate"
CONST_VALUE_MIN_TESTATBRATE = "minimum_tested_antibiotic_rate"
CONST_VALUE_MAX_TESTATBRATE = "maximum_tested_antibiotic_rate"
CONST_VALUE_MIN_RRATE = "minimum_resistant_rate"
CONST_VALUE_MAX_RRATE = "maximum_resistant_rate"
CONST_VALUE_MIN_IRATE = "minimum_intermediate_rate"
CONST_VALUE_MAX_IRATE = "maximum_intermediate_rate"
CONST_VALUE_MIN_SRATE = "minimum_susceptible_rate"
CONST_VALUE_MAX_SRATE = "maximum_susceptible_rate"


#INPUTs-OUTPUTs for AnnexC
CONST_FILENAME_WARD    = "dictionary_for_wards"
CONST_FILENAME_REPORT1 = AC.CONST_FILENAME_sec1_res_i 
CONST_FILENAME_HO_DEDUP= "AnnexC_dedup_profile"
CONST_FILENAME_ORIPARAM= ".\Programs\AMASS_amr\satscan_param.prm"
CONST_FILENAME_NEWPARAM= "satscan_param"
CONST_FILENAME_PROFILE = "profile_information"
CONST_FILENAME_AST = "ast_information" #v3.1 3102
CONST_FILENAME_LOCATION= "satscan_location"
CONST_FILENAME_RESULT  = "satscan_results"
CONST_FILENAME_INPUT   = "satscan_input"
CONST_FILENAME_ACLUSTER= "AnnexC_listofallclusters"
CONST_FILENAME_PCLUSTER= "AnnexC_listofpassedclusters"
CONST_FILENAME_AWARDPROF= "AnnexC_graphofallwardprof"
CONST_FILENAME_PWARDPROF= "AnnexC_graphofpassedwardprof"
CONST_FILENAME_COUNT   = "AnnexC_counts_by_organism"
CONST_FILENAME_MAIN_PDF     = "AnnexC_main_page"
CONST_FILENAME_SUPP_TCLUSTER= "AnnexC_supp_cluster_table"
CONST_FILENAME_SUPP_TPROFILE= "AnnexC_supp_profile_table"
CONST_FILENAME_SUPP_GRAPH   = "AnnexC_supp_wardprof_graph"
CONST_FILENAME_SUPP_PDF     = "AnnexC_supplementary_report"


#Report AnnexC and supplementary report style
CONST_STYLE_FONTB_OP = "<font color=\"#000080\">" #Blue
CONST_STYLE_FONTG_OP = "<font color=darkgreen>" #Green
CONST_STYLE_FONT_ED = "</font>"
CONST_STYLE_B_OP = "<b>" #Bold
CONST_STYLE_B_ED = "</b>"
CONST_STYLE_I_OP = "<i>" #Italic
CONST_STYLE_I_ED = "</i>"
CONST_STYLE_BREAKPARA = "<br/>"
CONST_STYLE_IDEN1_OP = "<para leftindent=\"35\">"
CONST_STYLE_IDEN2_OP = "<para leftindent=\"70\">"
CONST_STYLE_IDEN3_OP = "<para leftindent=\"105\">"
CONST_STYLE_IDEN_ED = "</para>"

#For calling, naming, and filtering pathogen of interest
#Able to add more interested profile for pathogens
#1 resistant profile  >>> "organism_escherichia_coli":[[CONST_NEWVARNAME_ASTCBPN,"1","crec"]]
#2 resistant profiles >>> "organism_escherichia_coli":[[CONST_NEWVARNAME_ASTCBPN,"1","crec"], [CONST_NEWVARNAME_AST3GC,"1",CONST_NEWVARNAME_ASTCBPN,"0","3gcr-csec"]]
#3 resistant profiles >>> "organism_escherichia_coli":[[AC.CONST_NEWVARNAME_AST3GC_RIS,"R",AC.CONST_NEWVARNAME_ASTCBPN_RIS,"S","3gcr-csec"]]
dict_ast = {"organism_staphylococcus_aureus":  [[AC.CONST_NEWVARNAME_ASTMRSA_RIS,"R","mrsa"]],
            "organism_enterococcus_faecalis":  [[CONST_NEWVARNAME_ASTVAN_RIS,    "R","vrefs"]],
            "organism_enterococcus_faecium":   [[CONST_NEWVARNAME_ASTVAN_RIS,    "R","vrefm"]],
            "organism_escherichia_coli":       [[AC.CONST_NEWVARNAME_ASTCBPN_RIS,"R","crec"]],
            "organism_klebsiella_pneumoniae":  [[AC.CONST_NEWVARNAME_ASTCBPN_RIS,"R","crkp"]],
            "organism_pseudomonas_aeruginosa": [[AC.CONST_NEWVARNAME_ASTCBPN_RIS,"R","crpa"]],
            "organism_acinetobacter_baumannii":[[AC.CONST_NEWVARNAME_ASTCBPN_RIS,"R","crab"]]}

#For namimg AST results
dict_ris = {"resistant":"R","intermediate":"I","susceptible":"S"}
#For calling, naming, and reporting pathogens
dict_org = {"mrsa"     :["organism_staphylococcus_aureus",  "Methicillin-resistant <i>S. aureus</i>",   "MRSA"],
            "vrefs"    :["organism_enterococcus_faecalis",  "Vancomycin-resistant <i>E. faecalis</i>",  "VREfs"], 
            "vrefm"    :["organism_enterococcus_faecium",   "Vancomycin-resistant <i>E. faecium</i>",   "VREfm"],
            "crec"     :["organism_escherichia_coli",       "Carbapenem-resistant <i>E. coli</i>",      "CREC"],
            "crkp"     :["organism_klebsiella_pneumoniae",  "Carbapenem-resistant <i>K. pneumoniae</i>","CRKP"],
            "crpa"     :["organism_pseudomonas_aeruginosa", "Carbapenem-resistant <i>P. aeruginosa</i>","CRPA"],
            "crab"     :["organism_acinetobacter_baumannii","Carbapenem-resistant <i>A. baumannii</i>", "CRAB"]}

#For calling and naming specimens
dict_spc = {"blo":"Blood specimen",
            "all":"All specimens"}
#For configuring parameters in SaTScan
dict_configuration_prm = {"CaseFile="                               :"satscan_input_", 
                          "CoordinatesFile="                        :"satscan_location_", 
                          "ResultsFile="                            :"satscan_results_", 
                          "CoordinatesType="                        :"0",
                          "AnalysisType="                           :"satscan_analysis_type",
                          "ModelType="                              :"satscan_model_type",
                          "ScanAreas="                              :"1",
                          "TimeAggregationUnits="                   :"3",
                          "MaxSpatialSizeInMaxCirclePopulationFile=":"satscan_maximum_spatial_size_of_total_isolate", 
                          "MaxSpatialSizeInDistanceFromCenter="     :"1",
                          "MaxTemporalSize="                        :"satscan_maximum_temporal_days",
                          "MinimumCasesInHighRateClusters="         :"satscan_minimum_isolates_in_high_clusters",
                          "MonteCarloReps="                         :"satscan_number_monte_carlo_replications"}

dict_configuration_prm_default = {"CaseFile="                               :"satscan_input_", 
                                  "CoordinatesFile="                        :"satscan_location_", 
                                  "ResultsFile="                            :"satscan_results_", 
                                  "CoordinatesType="                        :"0",
                                  "AnalysisType="                           :"3",
                                  "ModelType="                              :"9",
                                  "ScanAreas="                              :"1",
                                  "TimeAggregationUnits="                   :"3",
                                  "MaxSpatialSizeInMaxCirclePopulationFile=":"50", 
                                  "MaxSpatialSizeInDistanceFromCenter="     :"1",
                                  "MaxTemporalSize="                        :"100",
                                  "MinimumCasesInHighRateClusters="         :"2",
                                  "MonteCarloReps="                         :"9999"}

#For adding additional antibiotics for profiling
#Able to add more antibiotics for organisms
#1 antibiotic  >>> "organism_staphylococcus_aureus":[CONST_NEWVARNAME_PREFIX_RIS+"Cefoperazone_and_sulbactam"]
#2 antibiotics >>> "organism_staphylococcus_aureus":[CONST_NEWVARNAME_PREFIX_RIS+"Cefoperazone_and_sulbactam",CONST_NEWVARNAME_PREFIX_RIS+"Trimethoprim"]
# dict_configuration_astforprofile = {"organism_staphylococcus_aureus":[AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin", 
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Ofloxacin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Gentamicin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Amikacin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Sulfamethoxazole_and_trimethoprim",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Rifampicin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Teicoplanin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Daptomycin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Linezolid",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Ceftaroline",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"],
#                                     "organism_enterococcus_faecalis":[AC.CONST_NEWVARNAME_PREFIX_RIS+"Ciprofloxacin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Levofloxacin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin",
#                                                                       AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"],
#                                     "organism_enterococcus_faecium":[AC.CONST_NEWVARNAME_PREFIX_RIS+"Ciprofloxacin",
#                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Levofloxacin",
#                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin",
#                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"]}
##will be added in the future
dict_configuration_astforprofile = {"organism_staphylococcus_aureus":[[AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin", 
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Ofloxacin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Gentamicin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Amikacin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Sulfamethoxazole_and_trimethoprim",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Rifampicin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Teicoplanin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Daptomycin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Linezolid",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Ceftaroline",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"],
                                                                      ["Erythromycin",
                                                                       "Ofloxacin",
                                                                       "Gentamicin",
                                                                       "Amikacin",
                                                                       "Co-trimoxazole",
                                                                       "Rifampicin",
                                                                       "Teicoplanin",
                                                                       "Daptomycin",
                                                                       "Linezolid",
                                                                       "Ceftaroline",
                                                                       "Piperacillin/tazobactam"]],
                                    "organism_enterococcus_faecalis":[[AC.CONST_NEWVARNAME_PREFIX_RIS+"Ciprofloxacin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Levofloxacin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin",
                                                                      AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"],
                                                                      ["Ciprofloxacin",
                                                                       "Levofloxacin",
                                                                       "Erythromycin",
                                                                       "Piperacillin/tazobactam"]],
                                    "organism_enterococcus_faecium":[[AC.CONST_NEWVARNAME_PREFIX_RIS+"Ciprofloxacin",
                                                                     AC.CONST_NEWVARNAME_PREFIX_RIS+"Levofloxacin",
                                                                     AC.CONST_NEWVARNAME_PREFIX_RIS+"Erythromycin",
                                                                     AC.CONST_NEWVARNAME_PREFIX_RIS+"Piperacillin_and_tazobactam"],
                                                                     ["Ciprofloxacin",
                                                                      "Levofloxacin",
                                                                      "Erythromycin",
                                                                      "Piperacillin/tazobactam"]]}

#For setting criteria for selecting antibioitcs in profiling step
#i.e. select an antibiotic when 0.1<resistant_rate<99.9 >>>>> {"minimum_tested_antibiotic_rate":"0.1","maximum_tested_antibiotic_rate":"99.9"}
dict_configuration_profile = {CONST_VALUE_MIN_TESTATBRATE:"profiling_minimum_tested_isolate_for_antibiotic"}

dict_configuration_profile_default = {CONST_VALUE_MIN_TESTATBRATE:"90",
                                      CONST_VALUE_MAX_TESTATBRATE:"100",
                                      CONST_VALUE_MIN_RRATE      :"0.1",
                                      CONST_VALUE_MAX_RRATE      :"99.9",
                                      CONST_VALUE_MIN_SRATE      :"0.1",
                                      CONST_VALUE_MAX_SRATE      :"99.9"}

dict_intermediate_configtosatscan={"satscan_analysis_type":{"Purely Spatial"                      :"1", 
                                                            "Purely Temporal"                     :"2", 
                                                            "Retrospective Space-Time"            :"3", 
                                                            "Prospective Space-Time"              :"4",
                                                            "Spatial Variation in Temporal Trends":"5",
                                                            "Prospective Purely Temporal"         :"6",
                                                            "Seasonal Temporal"                   :"7"},
                                   "satscan_model_type":{"Discrete Poisson"      :"0", 
                                                         "Bernoulli"             :"1", 
                                                         "Space-Time Permutation":"2", 
                                                         "Ordinal"               :"3", 
                                                         "Exponential"           :"4", 
                                                         "Normal"                :"5", 
                                                         "Continuous Poisson"    :"6", 
                                                         "Multinomial"           :"7",
                                                         "Rank"                  :"8", 
                                                         "UniformTime"           :"9"},
                                   "satscan_scan_areas":{"High Rates":"1", 
                                                         "Low Rates" :"2", 
                                                         "Both Areas":"3"},
                                   "satscan_time_aggregation_units":{"None"   :"0", 
                                                                     "Year"   :"1",
                                                                     "Month"  :"2",
                                                                     "Day"    :"3", 
                                                                     "Generic":"4"}}
#Color palette for AST #v3.1 3102
dict_ast_color = {CONST_VALUE_NA          :"#000000",
                  CONST_VALUE_RESISTANT   :"#E69F00",
                  CONST_VALUE_INTERMEDIATE:"#FDF150",
                  CONST_VALUE_SUSCEPTIBLE :"#009E73"}