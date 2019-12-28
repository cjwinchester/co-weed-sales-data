import os
import glob
import time

import requests
from bs4 import BeautifulSoup

'''Fetches a current list of .xlsx files on the CO marijuana sales reports page and downloads the ones we don't have already.'''  # noqa

# get a list of files that already exist in the `raw-data` directory
raw_files = glob.glob(os.path.join('raw-data', '*.xlsx'))

# the URL of the page with the links to the monthly sales reports
url = 'https://www.colorado.gov/pacific/revenue/colorado-marijuana-sales-reports'  # noqa

# fetch the main page
r = requests.get(url)

# turn it into soup
soup = BeautifulSoup(r.text, 'html.parser')

# find all `a` tags with "xlsx" in the href attribute but not "CalendarReport"
sheets = soup.find_all('a', href=lambda x: x and 'xlsx' in x and 'CalendarReport' not in x)  # noqa

# list comprehension to get just the hrefs from that list
to_dl = [x['href'] for x in sheets]

# list comprehension to get the names of the files we already have
already_there = [x.split('/')[-1] for x in raw_files]

# loop over the list of spreadsheet links to download
for link in to_dl:

    # the name of the file to save locally is the last bit of the URL
    fname = link.split('/')[-1]

    # see if we've already got that file
    if fname not in already_there:

        # if not, join the destination directory to the filename
        dest = os.path.join('raw-data', fname)

        # print to let us know it's working
        print(dest)

        # fetch that file ...
        r = requests.get(link, stream=True)

        # and write it file on this computer
        with open(dest, 'wb') as f:

            # iterating a bit at a time over blocks of the file
            # http://docs.python-requests.org/en/master/api/#requests.Response.iter_content
            for block in r.iter_content(1024):
                f.write(block)

        # pause for 2 seconds before continuing
        time.sleep(2)
