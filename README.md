# OECD-web-scraping

This program allows users to launch a windowed application to launch the OECD website and land on the monthly LEI datasets and configure a CSV file to download.

When a user has this downloaded they can generate a diffusion index or multiple annual % change graphs. 

## User guide

### If using the non SDMX version

1. Launch OECD and download your desired dataset consisting of ALL countries you wish to use for either a diffusion index and or the annual % changes.
2. Once downloaded clicking either button will use the most recently downloaded CSV file from your downloads folder. Make sure you have not downloaded another in the meantime.
3. Once you have clicked either graph generating button check your downloads folder for the graphs outputted along with their dataset in the excel sheet

### If using the SDMX version

1. Go to line 47 to adjust the tickers/data you wish to obtain. A list of these can be found on the OECD site.
2. Start the program and click the query OECD button, this will take a few seconds and will automatically download the specified information you adjusted for above
3. Once downloaded you are free to generate the diffusion index or annual changes and are located in the local project directory alongside the queried OECD data.

