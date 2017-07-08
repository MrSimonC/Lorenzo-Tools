# Lorenzo Tools
A selection of python functions to help solve problems related to CSC Patient Administration System "Lorenzo".

## Files
- frequency_calculator.py
    - works out if a schedule pattern falls on a certain date (say, for example, a future bank holiday date)
- lorenzo_dcs.py
    - changes excel documents to .csv then loads into an MS SQL database, ready to be data quality checked
- lorenzo_dcs_vs_domain.py
    - compare dcs data vs domain data: loads data into database, runs stored procedure and outputs results to excel
- lorenzo_sp_from_spreadsheets.py
    - Go through a folder of xlsx files, collating data from each, outputting a single csv of all data
- lorenzo_xml_dtm_files.py
    - Analyse Lorenzo Letter (extracted) .DTM file and output used merge fields, output to clipboard