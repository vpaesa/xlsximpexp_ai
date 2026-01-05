#!/bin/bash

# This script depends on ssconvert (part of Gnumeric). You must install Gnumeric by yourself.
# Also depends on SQLite Linux shell (sqlite3). That is downloaded and unpacked by next two lines.
curl -C - --remote-name  https://sqlite.org/2025/sqlite-tools-linux-x64-3510100.zip
unzip -u sqlite-tools-linux-x64-3510100.zip sqlite3

echo -e "+++++++++++++++++++++++++++++++\nTesting xlsximport"
for llm in opus gemini copilot
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

echo -e "+++++++++++++++++++++++++++++++\nTesting xlsxexport"
for llm in opus gemini copilot opus_libxlsxwriter copilot_libxlsxwriter
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

# An 80MB spreadsheet, but full of non-compliances: it needs to be read by Libreoffice and saved again, then it becomes 127MB.
https://datacatalogfiles.worldbank.org/ddh-published/0037712/DR0095336/WDI_EXCEL_2025_12_19.zip

SELECT xlsx_import('WDIEXCEL.xlsx');

# This snippet generates the largest allowed number of rows in Excel.
./sqlite3 ':memory:' '.mode csv' '.headers on' '.once 14_headermillionrows_01.csv' "SELECT 'row' as header FROM generate_series(1, 1048575);"
ssconvert --import-type=Gnumeric_stf:stf_csvtab --export-type=Gnumeric_Excel:xlsx2 14_headermillionrows_01.csv 14_headermillionrows_01.xlsx

# make clean
rm expected_* importing_* exporting_*
