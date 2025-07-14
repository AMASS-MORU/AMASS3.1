#***-------------------------------------------------------------------------------------------------***#
#*** AutoMated tool for Antimicrobial resistance Surveillance System version 2.0 (AMASS version 2.0) ***#
#***-------------------------------------------------------------------------------------------------***#
# Aim: to enable hospitals with microbiology data available in electronic formats
# to report errors from all processes systematically.

# Created on 20th April 2022
import os #for getting file size
import AMASS_amr_const as AC
from pathlib import Path #for retrieving input's path
path = AC.CONST_PATH_ROOT
try:
    out_ = open(path + "error_log.txt","w")
    out_.write("Start of error_log.txt" + "\n")
    for filename in ["error_preprocess.txt","error_preprocess_whonet.txt","error_analysis_amr.txt","error_analysis_data.txt","error_report_amr.txt","error_report_data.txt","error_report_data_annexA.txt","error_logfile_amass.txt"]:
        if Path(filename).is_file():
            if os.path.getsize(filename) > 0:
                qc = open(filename,"r")
                qc_1 = qc.readline()
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("\t\t\t\t\t"+filename+"\t\t\t\t\t\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("--------------------------------------------------------------------------------------------------------------------\n")
                out_.write("\n")
                while qc_1 != "":
                    out_.write(qc_1)
                    qc_1 = qc.readline()
                qc.close()
        print (filename)
    out_.close()
except Exception:
    pass
