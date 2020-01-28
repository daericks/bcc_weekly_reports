# Weekly cannabis testing in California

## ETL
A script to clean up the data released weekly by the California BCC.
Separately download the data files (.xlsx) from the BCC. 
To run the script, navigate to `bcc_weekly_reports/etl` in your termainal run the script: `python bcc_etl.py`. This will ETL (extract, transform, load) the BCC data into 3 new summary (.csv) files. 

## Data Viz
I used this data to produce a couple of interactive visualizations using Tableau. Find them on my Tableau public: https://public.tableau.com/profile/david.erickson#!/

## Notebooks
The notebook has a step by step guide to ETL the data.


## Data Source
https://bcc.ca.gov/licensees/weekly_reports.html
