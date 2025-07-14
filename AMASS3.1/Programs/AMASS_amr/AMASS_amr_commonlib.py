#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 3.0 (AMASS version 3.0) ***#
#*** COMMON FUNCTION LIB                                                                             ***#
#***-------------------------------------------------------------------------------------------------***#
# @author: PRAPASS WANNAPINIJ
# Created on: 09 MAR 2023 
import pandas as pd #for creating and manipulating dataframe
import os
import sys
import csv
from xlsx2csv import Xlsx2csv
import logging #for creating error_log
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import AMASS_amr_const as AC
from pathlib import Path #for retrieving input's path
# Added build 3023 ---------------------------------------------------------------
def setlocalfont():
    LOCALFONT = "Helvetica"
    LOCALFONT_BOLD = "Helvetica-Bold"
    try:
        font_path = AC.CONST_PATH_CONFIG + "localfont.ttf"  
        if Path(font_path).is_file():
            LOCALFONT = "LOCALFONT"
            pdfmetrics.registerFont(TTFont(LOCALFONT, font_path))
    except:
        LOCALFONT = "Helvetica"
    try:
        font_path = AC.CONST_PATH_CONFIG + "localfont_bold.ttf"  
        if Path(font_path).is_file():
            LOCALFONT_BOLD = "LOCALFONT_BOLD"
            pdfmetrics.registerFont(TTFont(LOCALFONT_BOLD, font_path))
    except:
        LOCALFONT_BOLD = "Helvetica-Bold"
    return [LOCALFONT,LOCALFONT_BOLD]
# function for init log file
def initlogger(sprogname,slogname) :
    # Create a logging instance
    file_logger = logging.getLogger(sprogname)
    file_logger.setLevel(logging.INFO)
    # Assign a file-handler to that instance
    fh = logging.FileHandler(slogname)
    fh.setLevel(logging.INFO)
    # Format your logs (optional)
    fh.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    # Add the handler to your logging instance
    file_logger.addHandler(fh)
    return file_logger
# Print to console and log to log file   
def printlog(smessage,biserror,logger) :
    try:
        print(smessage)
        if biserror:
            logger.error(smessage)
        else:
            logger.info(smessage)
        return True
    except:
        return False
# read command line arguments
def getcmdargtodict(logger):
    dictarg = {}
    try:
        printlog("Command argument parameter is " + str(sys.argv), False, logger)
        listarg = list(str(sys.argv).strip("][").split(","))
        for sarg in listarg:
            li = list(sarg.replace("'","").split(":"))
            try:
                dictarg[str(li[0]).strip()] = str(li[1]).strip()
            except:
                pass
    except Exception as e: 
        printlog("Warning : unable to read command argument parameter", False, logger)
        logger.exception(e)
    #print(dictarg)
    return dictarg
# Check that is there either CSV or XLSX with the specific file name or not
def checkxlsorcsv(spath,sfilename) :
    bfound = False
    try:
        if os.path.exists(spath + sfilename + ".csv") :
            bfound = True
        else:
            if os.path.exists(spath + sfilename + ".xlsx") :
                bfound = True
    except:
        bfound = False
    return bfound
# Read csv or xlsx file, csv first, if no file convert xlsx to csv beflor load [WITH FORCE LOAD SOME COLUMNS as STRING]
def readxlsxorcsv_forcehntostring(spath,sfilename,hncol,logger) :
    df = pd.DataFrame()
    if hncol.strip() == "":
        df = readxlsxorcsv(spath,sfilename,logger)
    else:
        dict_converter = {hncol:str}
        try:
            df = pd.read_csv(spath + sfilename + ".csv",converters=dict_converter).fillna("")
        except:
            try:
                df= pd.read_csv(spath + sfilename + ".csv", encoding="windows-1252",converters=dict_converter).fillna("")
            except:
                try:
                    df= pd.read_excel(spath + sfilename + ".xlsx",converters=dict_converter).fillna("")
                except Exception as e:
                    printlog("Warning : using xlsxtocsv to convert, it may be strict open xml file format : "+ str(e),False,logger)
                    Xlsx2csv(spath + sfilename + ".xlsx", outputencoding="utf-8").convert(spath + sfilename + "_temp.csv")
                    df = pd.read_csv(spath + sfilename + "_temp.csv",converters=dict_converter).fillna("")
    return df
# Read csv or xlsx file, csv first, if no file convert xlsx to csv beflor load
def readxlsxorcsv(spath,sfilename,logger) :
    df = pd.DataFrame()
    try:
        df = pd.read_csv(spath + sfilename + ".csv").fillna("")
    except:
        try:
            df= pd.read_csv(spath + sfilename + ".csv", encoding="windows-1252").fillna("")
        except:
            try:
                df= pd.read_excel(spath + sfilename + ".xlsx").fillna("")
            except Exception as e:
                #printlog("Warning : using xlsxtocsv to convert, it may be strict open xml file format : "+ str(e),False,logger)
                Xlsx2csv(spath + sfilename + ".xlsx", outputencoding="utf-8").convert(spath + sfilename + "_temp.csv")
                df = pd.read_csv(spath + sfilename + "_temp.csv").fillna("")
    return df
# Read csv or xlsx file and specify header in columnheader list, csv first, if no file convert xlsx to csv beflor load
def readxlsorcsv_noheader(spath,sfilename,columnheader,logger) :
    ncol = len(columnheader)
    try:
        df = pd.read_csv(spath + sfilename + ".csv",header=None,names=columnheader,usecols=range(ncol)).fillna("")
    except:
        try:
            df= pd.read_csv(spath + sfilename + ".csv", encoding="windows-1252",header=None,names=columnheader,usecols=range(ncol)).fillna("")
        except:
            try:
                df= pd.read_excel(spath + sfilename + ".xlsx",header=None,names=columnheader,usecols=range(ncol)).fillna("")
            except Exception as e:
                #printlog("Warning : using xlsxtocsv to convert, it may be strict open xml file format : "+ str(e),False,logger)
                Xlsx2csv(spath + sfilename + ".xlsx", outputencoding="utf-8").convert(spath + sfilename + "_temp.csv")
                df = pd.read_csv(spath + sfilename + "_temp.csv",header=None,names=columnheader,usecols=range(ncol)).fillna("")
    return df
# Read csv or xlsx file and specify header in columnheader list, csv first, if no file convert xlsx to csv beflor load
def readxlsorcsv_noheader_forceencode(spath,sfilename,columnheader,sencoding,logger) :
    try:
        df = pd.read_csv(spath + sfilename + ".csv",header=None,names=columnheader,encoding=sencoding).fillna("")
    except:
        try:
            df= pd.read_excel(spath + sfilename + ".xlsx",header=None,names=columnheader).fillna("")
        except Exception as e:
            #printlog("Warning : using xlsxtocsv to convert, it may be strict open xml file format : "+ str(e),False,logger)
            Xlsx2csv(spath + sfilename + ".xlsx", outputencoding=sencoding).convert(spath + sfilename + "_temp.csv")
            df = pd.read_csv(spath + sfilename + "_temp.csv",header=None,names=columnheader,encoding=sencoding).fillna("")
    return df
# Save to csv
def fn_savecsv(df,fname,iquotemode,logger) :
    try:
        if iquotemode == 1:
            df.to_csv(fname,index=False, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        elif iquotemode == 2:
            df.to_csv(fname,index=False, quotechar='"', quoting=csv.QUOTE_ALL)  
        else:
            df.to_csv(fname,index=False)
        #logger.info("save csv file : " + fname)
        return True
    except Exception as e: # work on python 3.x
        printlog("Failed to save csv file : " + fname + " : "+ str(e),True,logger)
        return False
# Save to csv, with encoding
def fn_savecsvwithencoding(df,fname,sencoding,iquotemode,logger) :
    if sencoding.strip() == "":
        sencoding = 'utf-8'
    try:
        if iquotemode == 1:
            df.to_csv(fname,index=False, quotechar='"', encoding=sencoding, quoting=csv.QUOTE_NONNUMERIC)
        elif iquotemode == 2:
            df.to_csv(fname,index=False, quotechar='"', encoding=sencoding,quoting=csv.QUOTE_ALL)  
        else:
            df.to_csv(fname,index=False,encoding=sencoding)
        #logger.info("save csv file : " + fname)
        return True
    except Exception as e: # work on python 3.x
        printlog("Failed to save csv file : " + fname + " : "+ str(e),True,logger)
        return False
# Save to xlsx
def fn_savexlsx(df,fname,logger) :
    try:
        df.to_excel(fname,index=False)
        #logger.info("save xlsx file : " + fname)
        return True
    except Exception as e: # work on python 3.x
        printlog("Failed to save xlsx file : " + fname + " : "+ str(e),True,logger)
        return False
# Remove specific columns in dataframe
def fn_removecol(df,list_col) :
    list_colexist = []
    for scol in df.columns:
        if scol not in list_col:
            list_colexist.append(scol)    
    if len(list_colexist) > 0:
        return df[list_colexist]
    else:
        return df    
# Remove unused columns in dataframe 
def fn_keeponlycol(df,list_col) :
    list_colexist = []
    for scol in list_col:
        if scol in df.columns:
            list_colexist.append(scol)
    if len(list_colexist) > 0:
        return df[list_colexist]
    else:
        return df
# Get dict value with default
def fn_getdict(dict_p,skey,sdefault):
    sval = sdefault
    try:
        if skey in dict_p:
            sval = dict_p[skey]
        else:
            sval = sdefault
    except:
        sval = sdefault
    return sval
# Add antibiotic group base of antibiotic in group configure read from amr_const
def fn_addatbgroupbyconfig(df,dict_atbgroup,dict_ast,logger) :
    for satbg in dict_atbgroup:
        try:
            printlog("Note : Start determine RIS for antibiotic group : " +  satbg ,False,logger)
            oatbg = dict_atbgroup[satbg]
            satbgRIS = oatbg[0]
            #For test ------------------------
            #satbgRIS = satbgRIS +"_TEST"
            #---------------------------------
            df[satbgRIS] = ""
            latb = oatbg[1]
            dictR = {}
            dictI = {}
            dictS = {}
            bisempty = True
            #Buid dict for create condiftion
            for i in range(len(latb)):
                atb =  latb[i]
                #print(atb)
                if atb in df.columns:
                    dictR.update({atb: "R"})
                    dictI.update({atb: "I"})
                    dictS.update({atb: "S"})
                    bisempty = False
            if bisempty:
                printlog("Warning : No antibiotic columns for determine RIS for antibiotic group : " +  satbg ,False,logger)
            else:
                qR = ' | '.join(['({}=="{}")'.format(k, v) for k, v in dictR.items()])
                qI = ' | '.join(['({}=="{}")'.format(k, v) for k, v in dictI.items()])
                qS = ' | '.join(['({}=="{}")'.format(k, v) for k, v in dictS.items()])
                printlog("Note : Condition for determine RIS (R) for antibiotic group : " +  satbg + " : " + qR ,False,logger)
                df.loc[df.eval(qS), satbgRIS] = "S"
                df.loc[df.eval(qI), satbgRIS] = "I"
                df.loc[df.eval(qR), satbgRIS] = "R"
                df[satbg] = df[satbgRIS].map(dict_ast).fillna("NA") 
                df[satbgRIS] = df[satbgRIS].astype("category")
                df[satbg] = df[satbg].astype("category")
        except Exception as e:
            printlog("Warning : unable to determine RIS for antibiotic group : " +  satbg + " : "+ str(e),False,logger)   
    return df
   
# Change field type to category to save mem space
def fn_df_tocategory_datatype(df,list_col,logger) :
    curcol = ""
    for scol in list_col:
        curcol = scol
        if scol in df.columns:
            try:
                df[scol] = df[scol].astype("category")
                printlog("Note : convert field type to category for column : " + curcol,False,logger)  
            except Exception as e: # work on python 3.x
                printlog("Warning : convert field type to category type : " + curcol+ " : "+ str(e),False,logger)         
    return df
# Trim space and unreadable charecters
def fn_df_strstrips(df,list_col,logger) :
    curcol = ""
    for scol in list_col:
        curcol = scol
        if scol in df.columns:
            try:
                df[scol] = df[scol].astype("string").str.strip()
                printlog("Note : trimmed column : " + curcol,False,logger)  
            except Exception as e: # work on python 3.x
                printlog("Warning : unable to trim data in column : " + curcol+ " : "+ str(e),False,logger)         
    return df
def fn_clean_date(df,oldfield,cleanfield,dformat,logger):
    printlog("Note : Converting data field: " + oldfield + " to " + cleanfield,False,logger)  
    CDATEFORMAT_YMD =["%Y/%m/%d","%y/%m/%d","%Y%m%d"] 
    CDATEFORMAT_DMY =["%d/%m/%Y","%d/%m/%y","%d%m%Y"]
    CDATEFORMAT_MDY =["%m/%d/%Y","%m/%d/%y","%m%d%Y"]
    CDATEFORMAT_OTH =["%d/%b/%Y","%d/%b/%y","%d/%B/%Y","%d/%B/%y"]
    cleanfieldtemp = cleanfield + "_tmpamassf"
    cleanfieldtemp2 = cleanfield + "_tmpamassf2"
    cft_1 = cleanfield + "_d1"
    cft_2 = cleanfield + "_d2"
    cft_3 = cleanfield + "_d3"
    if oldfield != cleanfield:
        try:
            isalreadydatecol = pd.api.types.is_datetime64_any_dtype(df[oldfield])
        except:
            isalreadydatecol = False
        if isalreadydatecol != True:
            df[cleanfield] = df[oldfield].astype("string").str.replace(r'\.0$','',regex=True)
            df[cleanfield] = df[cleanfield].fillna("1900-01-01")
            df[cleanfield] = df[cleanfield].str.rsplit(" ", n = 1, expand = True)[0]
            iDMY = 0
            iYMD = 0
            iMDY = 0
            iOTH = 0
            try:
                df[cleanfield] = df[cleanfield].str.replace('-', '/', regex=False)
                df[cft_1] = df[cleanfield].str.split("/", n = 2, expand = True)[0]
                df[cft_1] = pd.to_numeric(df[cft_1],downcast='signed',errors='coerce')
                df[cft_2] = df[cleanfield].str.split("/", n = 2, expand = True)[1]
                df[cft_2] = pd.to_numeric(df[cft_2],downcast='signed',errors='coerce')
                iDMY = len(df[(df[cft_1]>12) & (df[cft_1]<32)])                   
                iYMD = len(df[(df[cft_1]>31)])
                iMDY = len(df[(df[cft_2]>12) & (df[cft_2]<32)])
                df = df.drop(columns=[cft_1])  
                df = df.drop(columns=[cft_2]) 
                #print('Count date format DMY:' + str(iDMY) + ', MDY:' + str(iMDY) + ', YMD:' + str(iYMD))
                #printlog("Able to specify string date format of " + oldfield,False,logger)  
            except Exception as e:
                printlog("Warning string date format of " + oldfield + " may be not in convert format defined or in other format", False, logger)
                #logger.exception(e)
                iOTH = 1
            df_format = pd.DataFrame({'fname':['YMD','DMY','MDY','Others'],'fcount':[iYMD,iDMY,iMDY,iOTH],'cformat':[CDATEFORMAT_YMD,CDATEFORMAT_DMY,CDATEFORMAT_MDY,CDATEFORMAT_OTH]}) 
            df_format = df_format.sort_values(by=['fcount'],ascending=False)
            df[cleanfieldtemp] = df[cleanfield]
            bfirstformat = True
            idf_rt = 0
            try:
                idf_rt = len(df)
                iprevconverted = 0
                printlog("Total rows : " + str(idf_rt),False,logger) 
            except:
                pass
            for index, row in df_format.iterrows():
                #print('Convert data format: ' + row['fname'] )
                for sf in row['cformat']:
                    #print(sf)
                    if bfirstformat:
                        df[cleanfield] = pd.to_datetime(df[cleanfield], format=sf, errors="coerce")
                        #print(df[cleanfield])
                    else:
                        if df[cleanfield].isnull().values.any() == False:
                            break
                        df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[cleanfieldtemp], format=sf, errors="coerce"))
                        #print(df[cleanfield])
                    bfirstformat = False
                    try:
                        icurna = df[cleanfield].isnull().sum()
                        icurconverted = (idf_rt - icurna) - iprevconverted
                        printlog("try format : " + sf + " : converted : " + str(icurconverted),False,logger) 
                        iprevconverted= idf_rt - icurna
                    except:
                        pass
            #Build 3027 - add try convert format of yyyyMMddhhnnss or ddMMyyyyhhnnss or MMddyyyy
            if df[cleanfield].isnull().values.any() == True:
                printlog("Trim data string for first 8 charector and try format : yyyyMMdd*, ddMMyyyy*, MMddyyyy* ",False,logger) 
                try:
                    df[cleanfieldtemp2] = df[oldfield].astype("string").str.replace(r'\.0$','',regex=True).str[:8]
                    df[cft_1]  = pd.to_numeric(df[cleanfieldtemp2].str[:2],downcast='signed',errors='coerce').fillna(0)
                    df[cft_2]  = pd.to_numeric(df[cleanfieldtemp2].str[2:4],downcast='signed',errors='coerce').fillna(0)
                    df[cft_3]  = pd.to_numeric(df[cleanfieldtemp2].str[:4],downcast='signed',errors='coerce').fillna(0)
                    iDMY = len(df[(df[cft_1]>12) & (df[cft_1]<32) & ~((df[cft_3]>2000) & (df[cft_3]<2100))])                   
                    iYMD = len(df[(df[cft_3]>2000) & (df[cft_3]<2100)])   
                    iMDY = len(df[(df[cft_2]>12) & (df[cft_2]<32) & ~((df[cft_3]>2000) & (df[cft_3]<2100))]) 
                    df_format = pd.DataFrame({'fname':['YMD','DMY','MDY'],'fcount':[iYMD,iDMY,iMDY],'cformat':["%Y%m%d","%d%m%Y","%m%d%Y"]}) 
                    df_format = df_format.sort_values(by=['fcount'],ascending=False)
                    for index, row in df_format.iterrows():
                        iprevna = 0
                        try:
                            iprevna = df[cleanfield].isnull().sum()
                            df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[cleanfieldtemp2], format=row['cformat'], errors="coerce"))
                            icurconverted = iprevna - df[cleanfield].isnull().sum()
                            printlog("try format : " + row['cformat'] + " : converted : " + str(icurconverted),False,logger) 
                        except:
                            printlog("warning : error try format : " + row['cformat'],False,logger) 
                    df = df.drop(columns=[cft_1])  
                    df = df.drop(columns=[cft_2]) 
                    df = df.drop(columns=[cft_3]) 
                    df = df.drop(columns=[cleanfieldtemp2]) 
                except:
                    printlog("warning : error try format : yyyyMMdd*, ddMMyyyy*, MMddyyyy* ",False,logger) 
            if df[cleanfield].isnull().values.any() == True:
                df[cleanfield] = df[cleanfield].fillna(pd.to_datetime(df[oldfield], errors="coerce"))
                printlog("Try pandas general date convertion",False,logger) 
            #df.loc[df[cleanfield]<datetime(1900, 1, 1),cleanfield] = np.nan
            df = df.drop(columns=[cleanfieldtemp])  
        else:
            df[cleanfield] = df[oldfield]
            printlog("Note: " + oldfield + " is already date time data type",False,logger) 
    return df