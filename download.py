import csv
from io import StringIO

import requests

sheet_id = '1br_cwfHy24d2R2bcXacb2KarOIBKGrbR'
csv_link = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'  # noqa
csv_outfile = 'co-cannabis-sales.csv'

headers = [
    'monthyear',
    'month',
    'year',
    'county',
    'amount',
    'sales_type'
]

r = requests.get(csv_link)

data = list(csv.reader(StringIO(r.text)))[5:-10]

monthly_totals = {}

with open(csv_outfile, 'w') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(headers)

    for row in data:
        month, year, county, medical, retail = [x.strip().upper() for x in row]
        year_month = f'{year}{month.zfill(2)}'

        if not monthly_totals.get(year_month):
            monthly_totals[year_month] = {
                'total_medical': 0,
                'non_nr_total_medical': 0,
                'total_retail': 0,
                'non_nr_total_retail': 0
            }

        try:
            medical = int(medical.replace('$', '').replace(',', ''))
        except ValueError:
            pass

        try:
            retail = int(retail.replace('$', '').replace(',', ''))
        except ValueError:
            pass

        if county == 'TOTAL':
            monthly_totals[year_month]['total_medical'] = medical
            monthly_totals[year_month]['total_retail'] = retail
        else:
            if type(medical) == int:
                monthly_totals[year_month]['non_nr_total_medical'] += medical
                data_med = [
                    year_month,
                    month,
                    year,
                    county,
                    medical,
                    'medical'
                ]
                writer.writerow(data_med)

            if type(retail) == int:
                monthly_totals[year_month]['non_nr_total_retail'] += retail
                data_ret = [
                    year_month,
                    month,
                    year,
                    county,
                    retail,
                    'retail'
                ]
                writer.writerow(data_ret)

    for monthyear in monthly_totals:
        totals = monthly_totals[monthyear]
        nr_med = totals['total_medical'] - totals['non_nr_total_medical']
        nr_ret = totals['total_retail'] - totals['non_nr_total_retail']
        year = int(monthyear[:4])
        month = int(monthyear[-2:])

        writer.writerow([monthyear, month, year, 'SUM OF NR COUNTIES', nr_ret, 'retail'])
        writer.writerow([monthyear, month, year, 'SUM OF NR COUNTIES', nr_med, 'medical'])
