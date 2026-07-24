MDR-DHF Organizer v5 - audited for Windows PowerShell 5.1

PREVIEW ONLY:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\organize_mdr_dhf.ps1"

EXECUTE AFTER REVIEWING BOTH CSV FILES:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\organize_mdr_dhf.ps1" -Execute

SEARCH THE MISSING CSV AGAIN BY DOCUMENT NAME (PREVIEW):
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\organize_mdr_dhf.ps1" -SearchMissingByName -ListPath ".\DHF_未找到编号_20260721_170030.csv"

COPY REVIEWED NAME MATCHES TO DESKTOP\DHF_undef:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\organize_mdr_dhf.ps1" -SearchMissingByName -ListPath ".\DHF_未找到编号_20260721_170030.csv" -Execute

PREVIEW DHF CLASSIFICATION BY PROJECT STAGE:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\categorize_dhf_by_stage.ps1"

COPY DESKTOP\DHF INTO DESKTOP\DHF_categorized\T1&2,T3,T4,T5,T6:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\categorize_dhf_by_stage.ps1" -Execute

BATCH FORMAT ALL WORD FILES UNDER DHF_categorized_bk:
Import "BatchFormatCategorizedDocs.bas" in Word VBA and run BatchFormatCategorizedWordDocs.

The preview mode exits before Copy-Item. No Move-Item or Remove-Item is used.
