@ECHO OFF

del ".\Report_with_patient_identifiers\*.xlsx"
del ".\Report_with_patient_identifiers\*.csv"
del ".\Report_with_patient_identifiers\*.pdf"
del ".\ResultData\*.xlsx"
del ".\ResultData\*.csv"
del ".\ResultData\*.png"
del ".\Temporary_folder_with_patient_identifiers\*.txt"
del ".\Temporary_folder_with_patient_identifiers\*.png"
del ".\Temporary_folder_with_patient_identifiers\*.csv"
del ".\Temporary_folder_with_patient_identifiers\*.xlsx"
del ".\Temporary_folder_with_patient_identifiers\*.prm"
del /q ".\Temporary_folder_with_patient_identifiers\*.*"
del ".\Variables\*.csv"

del ".\AMR_surveillance_report.pdf"
del ".\microbiology_data_reformatted.xlsx"
del ".\data_verification_logfile_report.pdf"
del ".\error_log.txt"
del ".\log_amr_analysis.txt"
del ".\log_dataverification_log.txt"

rmdir ".\Temporary_folder_with_patient_identifiers\"
rmdir ".\Report_with_patient_identifiers\"
rmdir ".\ResultData\"
rmdir ".\Variables\"
mkdir Temporary_folder_with_patient_identifiers
mkdir Report_with_patient_identifiers
mkdir ResultData
mkdir Variables

echo Start Preprocessing: %date% %time%
.\Programs\Python-Portable\Portable_Python-3.8.9\App\Python\python.exe -W ignore .\Programs\AMASS_preprocess\AMASS_preprocess_version_2.py
.\Programs\Python-Portable\Portable_Python-3.8.9\App\Python\python.exe -W ignore .\Programs\AMASS_preprocess\AMASS_preprocess_whonet_version_2.py
echo Start AMR analysis
.\Programs\Python-Portable\Portable_Python-3.8.9\App\Python\python.exe -W ignore .\Programs\AMASS_amr\AMASS3.0_amr_analysis.py
echo Start generating "Data verificator logfile report"
.\Programs\Python-Portable\Portable_Python-3.8.9\App\Python\python.exe -W ignore .\Programs\AMASS_amr\AMASS_logfile_version_2.py
.\Programs\Python-Portable\Portable_Python-3.8.9\App\Python\python.exe -W ignore .\Programs\AMASS_amr\AMASS_logfile_err_version_2.py
copy ".\Configuration\Configuration.xlsx" ".\ResultData\Configuration_used.xlsx"
del ".\ResultData\logfile_age.xlsx"
del ".\ResultData\logfile_ast.xlsx"
del ".\ResultData\logfile_discharge.xlsx"
del ".\ResultData\logfile_gender.xlsx"
del ".\ResultData\logfile_organism.xlsx"
del ".\ResultData\logfile_specimen.xlsx"
del ".\ResultData\*.png"
del /q ".\Temporary_folder_with_patient_identifiers\*.*"
rmdir ".\Temporary_folder_with_patient_identifiers\"
del ".\error_analysis*.txt"
del ".\error_report*.txt"
del ".\error_logfile_amass.txt"
del ".\error_preprocess.txt"
del ".\error_preprocess_whonet.txt"
del ".\Report_with_patient_identifiers\Report_with_patient_identifiers_annexA_withstatus.xlsx"
del ".\Report_with_patient_identifiers\Report_with_patient_identifiers_annexB_withstatus.xlsx"
echo Finish running AMASS: %date% %time%

