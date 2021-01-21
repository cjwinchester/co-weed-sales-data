import os
import glob
import time
from datetime import datetime
import string

import requests
from bs4 import BeautifulSoup

'''Fetches a current list of .xlsx files on the CO marijuana sales reports page (now saved as google drive links because why not) and downloads the ones we don't have already.'''  # noqa

# folder to hold all the things
raw_file_dir = 'raw-data'

# get a list of xlsx files we already have
raw_files = glob.glob(os.path.join(raw_file_dir, '*.xlsx'))

# the URL to the page with links to the monthly sales reports
url = 'https://www.colorado.gov/pacific/revenue/colorado-marijuana-sales-reports'  # noqa

# drive file download pattern via https://stackoverflow.com/a/39087286
drive_download_pattern = 'https://drive.google.com/uc?export=download&id={drive_file_id}'  # noqa

# fetch the main page
r = requests.get(url)

# turn it into soup
soup = BeautifulSoup(r.text, 'html.parser')

# grab the description lists
dls = soup.find_all('dl')

# jfc people, did we do this in dreamweaver
for dl in dls:

    # grab the "description terms" and the
    # "description definitions" 🙄
    dts = dl.find_all('dt')
    dds = dl.find_all('dd')

    # have to do it this way because two years might be
    # lumped together into one section, why not
    zipped = list(zip(dts, dds))

    for pair in zipped:

        # the year
        year = pair[0].get_text(strip=True)

        # the month markup
        months = pair[1]

        # sure cool let's just dump everything into a
        # single graf with <br/>s everywhere
        months_split = str(months.p).split('<br/>')

        # go through each of these dumbshits
        for month in months_split:

            # get the name of the month
            month_name = month.split('<a')[0].split('>')[-1].strip()

            # kill nonprintable cruft in some strings
            month_clean = ''.join(
                [x for x in month_name if x in string.printable]
            )

            # grab the 0-padded month number
            month_num = datetime.strptime(
                month_clean, '%B'
            ).strftime('%m')

            # turn this chunk into soup
            month_soup = BeautifulSoup(month, 'html.parser')

            # find the "Excel" link
            drive_link = month_soup.find('a', text='Excel')['href']

            # extract the drive ID
            drive_id = drive_link.split('/d/')[-1].split('/')[0]

            # build the file path
            filename = f'{year}{month_num}_raw.xlsx'
            filepath = os.path.join(raw_file_dir, filename)

            # skip if we already have it
            if filepath not in raw_files:

                # otherwise download
                print(f'Downloading {filename} ...')

                drive_url = drive_download_pattern.format(drive_file_id=drive_id)  # noqa

                r = requests.get(drive_url, stream=True)

                with open(filepath, 'wb') as outfile:
                    for block in r.iter_content(1024):
                        outfile.write(block)

                # pause for a sec before continuing
                time.sleep(1)
