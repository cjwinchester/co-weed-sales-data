# Colorado cannabis sales reports
Some Python code to fetch, process and analyze [cannabis sales reports data in Colorado](https://www.colorado.gov/pacific/revenue/colorado-marijuana-sales-reports).

**Note**: The state stopped producing individual monthly XLSX files in favor of [a unified Google Sheet](https://docs.google.com/spreadsheets/d/1br_cwfHy24d2R2bcXacb2KarOIBKGrbR/edit#gid=1659782909) sometime in 2021/22.

## Run the code

- Clone or download this repo
- `cd` into the directory
- Create & activate a virtual environment with your tooling of choice
- `pip install -r requirements.txt`
- To download the data and write to local file: `python download.py`
- To run the notebook server: `jupyter notebook`