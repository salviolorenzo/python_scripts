import xlrd
import json
import datetime


def xls_to_json(file):

    if file.slice(-1, -4, 1) == 'txt':
        print('txt')

        x = xlrd.open_workbook('./DailyReports/' + file)
        x1 = x.sheet_by_name('Sheet1')
        # writecsv = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
        data = {}
        data['rows'] = []
        date = 0
        for rownum in range(x1.nrows):  # To determine the total rows.
            row = x1.row_values(rownum)
            if rownum == 0:
                pass
            elif row[0] == 'Date':
                print(date)
                date = date + 1
            else:
                data['rows'].append({
                    'date': date,
                    'ticket': row[0],
                    'worked_by': row[1],
                    'status': row[2],
                    'tested_by': row[3]
                })
                # print(rowObject)
                # writecsv.writerow(rowObject)
        # csvfile.close()
        # print(data)
        return data


# xls_to_json('Dedicated-Daily-Report_02_19.xlsx')
with open('data.json', 'w') as outfile:
    json.dump(xls_to_json('Dedicated-Daily-Report_02_19.txt'), outfile)


# data = {}
# data['people'] = []
# data['people'].append({
#     'name': 'Scott',
#     'website': 'stackabuse.com',
#     'from': 'Nebraska'
# })
# data['people'].append({
#     'name': 'Larry',
#     'website': 'google.com',
#     'from': 'Michigan'
# })
# data['people'].append({
#     'name': 'Tim',
#     'website': 'apple.com',
#     'from': 'Alabama'
# })

# with open('data.txt', 'w') as outfile:
#     json.dump(data, outfile)
