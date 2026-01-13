#!/bin/bash

# This script depends on ssconvert (part of Gnumeric). You must install Gnumeric by yourself.
# Also depends on SQLite Linux shell (sqlite3). That is downloaded and unpacked by next two lines.
#curl -C - --remote-name  https://sqlite.org/2026/sqlite-tools-linux-x64-3510200.zip
#unzip -u sqlite-tools-linux-x64-3510200.zip sqlite3

echo -e "+++++++++++++++++++++++++++++++\nMinimal Testing of xlsximport and xlsxexport"
for llm in chatgpt chatgpt_libxlsxwriter gemini gemini_libxlsxwriter opus opus_libxlsxwriter copilot copilot_libxlsxwriter
do
  echo -e "-------------------------------\nLLM: ${llm}"
  ./sqlite3 ':memory:' '.mode box' ".load ../${llm%_libxlsxwriter}/xlsximport.so" "SELECT xlsx_import('09_severalsheets_t_06.xlsx');" \
  '.schema' 'select * from "00";' "SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('09_severalsheets_t_06.xlsx');" 'SELECT xlsx_import_version();' \
  "SELECT * FROM sqlite_master WHERE type='table';" ".load ../${llm}/xlsxexport.so" "SELECT xlsx_export('validating_09_severalsheets_t_06.xlsx');" \
  'SELECT xlsx_export_version();'
done

echo -e "+++++++++++++++++++++++++++++++\nThorough Testing xlsximport"
for llm in opus gemini copilot chatgpt
do
  echo -e "-------------------------------\nLLM: ${llm}"
  for i in ??_*_??.xlsx
  do
    testcase=${i%.xlsx}
    sheets=${i: -7:2}
    echo "Testcase ${i} has ${sheets} sheets"
    ssconvert --import-type=Gnumeric_Excel:xlsx --export-type=Gnumeric_stf:stf_csv --export-file-per-sheet $i expected_${testcase}.csv
    for expected_sheet in expected_${testcase}.csv.*
    do
      echo "Testing ${i} sheet ${sheetid}"
      importing_sheet="importing_${expected_sheet#expected_}"
      sheetid=$(printf "%02d" "${expected_sheet##*.}")
      # '.trace'  
      ./sqlite3 ':memory:' '.mode csv' '.headers on' ".import ${expected_sheet} ThisIsWhatIExpect" ".once ${expected_sheet}" "SELECT * from ThisIsWhatIExpect;" ".load ../${llm}/xlsximport.so" "SELECT xlsx_import('$i');" ".once ${importing_sheet}" "SELECT * from \"${sheetid}\";" 
      cmp $expected_sheet $importing_sheet
      if [ $? -eq 0 ]
      then echo "Passed ${testcase} sheet ${sheetid}"
      else echo "Failed ${testcase} sheet ${sheetid}"
      fi
      echo
    done  
  done
done

echo -e "+++++++++++++++++++++++++++++++\nThorough Testing xlsxexport"
for llm in chatgpt chatgpt_libxlsxwriter gemini gemini_libxlsxwriter opus opus_libxlsxwriter copilot copilot_libxlsxwriter
do
  echo -e "-------------------------------\nLLM: ${llm}"
  #for i in ??_*_??.xlsx
  for i in 00_headertworows_01.xlsx 14_headermillionrows_01.xlsx
  do
    testcase=${i%.xlsx}
    sheets=${i: -7:2}
    echo "Testcase ${i} has ${sheets} sheets"
    #ssconvert --import-type=Gnumeric_Excel:xlsx --export-type=Gnumeric_stf:stf_csv --export-file-per-sheet $i expected_${testcase}.csv
    for expected_sheet in expected_${testcase}.csv.*
    do
      echo "Testing ${i} sheet ${sheetid}"
      exporting_sheet="exporting_${expected_sheet#expected_}"
      sheetid=$(printf "%02d" "${expected_sheet##*.}")
      # '.trace'  
      ./sqlite3 ':memory:' '.mode csv' '.headers on' ".import ${expected_sheet} \"${sheetid}\"" ".load ../${llm}/xlsxexport.so" "SELECT xlsx_export('exporting_$i', '${sheetid}');"
      ssconvert --import-type=Gnumeric_Excel:xlsx --export-type=Gnumeric_stf:stf_csv exporting_$i $exporting_sheet
      ./sqlite3 ':memory:' '.mode csv' '.headers on' ".import ${exporting_sheet} ThisIsWhatIExport" ".once ${exporting_sheet}" "SELECT * from ThisIsWhatIExport;"
      cmp $expected_sheet $exporting_sheet
      if [ $? -eq 0 ]
      then echo "Passed ${testcase} sheet ${sheetid}"
      else echo "Failed ${testcase} sheet ${sheetid}"
      fi
      echo
    done  
  done
done
exit

echo -e "+++++++++++++++++++++++++++++++\nTest with a 117MB spreadsheet of xlsximport and xlsxexport"
ls -l ../attic/WDIEXCEL_Libreoffice.xlsx
for llm in chatgpt chatgpt_libxlsxwriter gemini gemini_libxlsxwriter opus opus_libxlsxwriter copilot copilot_libxlsxwriter
do
  echo -e "-------------------------------\nLLM: ${llm}"
  ./sqlite3 ':memory:' ".load ../${llm%_libxlsxwriter}/xlsximport.so" '.output counts.sql' \
  "SELECT 'select ' || quote(sheet_name) || ' as sheet_name, count(*) from \"' || sheet_name || '\" union all ' FROM xlsx_import_sheetnames('../attic/WDIEXCEL_Libreoffice.xlsx');" \
  ".print \"select '', '';\""
  time ./sqlite3 ':memory:' '.mode box' ".load ../${llm%_libxlsxwriter}/xlsximport.so" "SELECT xlsx_import('../attic/WDIEXCEL_Libreoffice.xlsx');" \
  '.read counts.sql' 'SELECT xlsx_import_version();' '.schema' \
  ".load ../${llm}/xlsxexport.so" "SELECT xlsx_export('validating_WDIEXCEL_Libreoffice_${llm}.xlsx');" 'SELECT xlsx_export_version();'
done
ls -l validating_WDIEXCEL_Libreoffice_*.xlsx

exit

# Some test snippets

# https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
soffice --headless --convert-to "xlsx:Calc MS Excel 2007 XML" WDIEXCEL_other.xlsx
soffice --headless --convert-to "xlsx:Calc Office Open XML" WDIEXCEL_other.xlsx
soffice --headless --convert-to "xlsx:Calc Office Open XML" WDIEXCEL_other.xlsx

# Minimal test of xlsx_import, xlsx_import_sheetnames, xlsx_import_version
./sqlite3 ':memory:' '.mode box' '.load ../copilot/xlsximport.so' "SELECT xlsx_import('09_severalsheets_t_06.xlsx');" '.schema' 'select * from "00";' "SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('09_severalsheets_t_06.xlsx');"
./sqlite3 ':memory:' '.mode box' '.load ../copilot/xlsximport.so' "SELECT xlsx_import('09_severalsheets_t_06.xlsx');" '.schema' 'select * from "00";' "SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('09_severalsheets_t_06.xlsx');"

# Minimal test of xlsx_import, xlsx_import_sheetnames, xlsx_import_version, xlsx_export, xlsx_export_version
./sqlite3 ':memory:' '.mode box' '.load ../gemini/xlsximport.so' "SELECT xlsx_import('09_severalsheets_t_06.xlsx');" '.schema' 'select * from "00";' "SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('09_severalsheets_t_06.xlsx');" 'SELECT xlsx_import_version();' "SELECT * FROM sqlite_master WHERE type='table';" '.load ../gemini/xlsxexport.so' "SELECT xlsx_export('validating_09_severalsheets_t_06.xlsx');" 'SELECT xlsx_export_version();'

# An 80MB spreadsheet, but full of non-compliances: it needs to be read by Libreoffice and saved again, then it becomes 117MB.
https://datacatalogfiles.worldbank.org/ddh-published/0037712/DR0095336/WDI_EXCEL_2025_12_19.zip

# Test with a 117MB spreadsheet
for llm in opus gemini copilot chatgpt opus_libxlsxwriter gemini_libxlsxwriter copilot_libxlsxwriter chatgpt_libxlsxwriter
do
  echo -e "-------------------------------\nLLM: ${llm}"
  ./sqlite3 ':memory:' '.mode box' '.timer on' ".load ../${llm%_libxlsxwriter}/xlsximport.so" "SELECT xlsx_import('../attic/WDIEXCEL_Libreoffice.xlsx');" \
  "SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('../attic/WDIEXCEL_Libreoffice.xlsx');" 'SELECT xlsx_import_version();' \
  ".load ../${llm}/xlsxexport.so" "SELECT xlsx_export('validating_WDIEXCEL_Libreoffice_${llm}.xlsx');" 'SELECT xlsx_export_version();'
done

SELECT xlsx_import('WDIEXCEL.xlsx');

# This snippet generates the largest allowed number of rows in Excel.
./sqlite3 ':memory:' '.mode csv' '.headers on' '.once 14_headermillionrows_01.csv' "SELECT 'row' as header FROM generate_series(1, 1048575);"
ssconvert --import-type=Gnumeric_stf:stf_csvtab --export-type=Gnumeric_Excel:xlsx2 14_headermillionrows_01.csv 14_headermillionrows_01.xlsx

# make clean
rm expected_* importing_* exporting_*

# Check the number of elements in the XML file
unzip -p validating_H2S_gemini.xlsx xl/worksheets/sheet1.xml|xmlstarlet el|sort|uniq -c
unzip -p validating_H2S_opus.xlsx xl/worksheets/sheet1.xml|xmlstarlet el|sort|uniq -c
unzip -p validating_H2S_gemini_libxlsxwriter.xlsx xl/worksheets/sheet1.xml|xmlstarlet el|sort|uniq -c

# Canonicalize the XML file
unzip -p validating_H2S_gemini.xlsx xl/styles.xml|xmllint --c14n11 - | xmllint --format - > validating_H2S_gemini_canonicalized_styles.xml
unzip -p validating_H2S_gemini_libxlsxwriter.xlsx xl/styles.xml|xmllint --c14n11 - | xmllint --format - > validating_H2S_gemini_libxlsxwriter_canonicalized_styles.xml
unzip -p validating_H2S_opus.xlsx xl/styles.xml|xmllint --c14n11 - | xmllint --format - > validating_H2S_opus_canonicalized_styles.xml

unzip -p validating_H2S_gemini.xlsx '_rels/.rels' | xmllint --format -
unzip -p validating_H2S_gemini_libxlsxwriter.xlsx '_rels/.rels'| xmllint --format -
unzip -p validating_H2S_opus.xlsx '_rels/.rels'| xmllint --format -

unzip -p validating_H2S_gemini.xlsx xl/workbook.xml | xmllint --format -
unzip -p validating_H2S_gemini_libxlsxwriter.xlsx xl/workbook.xml | xmllint --format -
unzip -p validating_H2S_opus.xlsx xl/workbook.xml | xmllint --format -

unzip -p validating_H2S_gemini.xlsx xl/_rels/workbook.xml.rels | xmllint --format -
unzip -p validating_H2S_gemini_libxlsxwriter.xlsx xl/_rels/workbook.xml.rels | xmllint --format -
unzip -p validating_H2S_opus.xlsx xl/_rels/workbook.xml.rels | xmllint --format -

# Test spreadsheets (XLS and XLSX)
https://github.com/tidyverse/readxl/tree/main/tests/testthat/sheets

# Example
https://github.com/tidyverse/readxl/tree/main/inst/extdata
