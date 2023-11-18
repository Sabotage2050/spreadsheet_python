#!python3

import openpyxl, pprint
print('Opening workbook ...')

wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb['Population by Census Tract']
countryData = {}
print('Reading rows...')
for a in range(2, sheet.max_row + 1):
    state = sheet['B'+str(a)].value
    country = sheet['C'+str(a)].value
    pop = sheet['D' + str(a)].value
    countryData.setdefault(state, {})
    countryData[state].setdefault(country, {'tracts': 0, 'pop': 0})
    countryData[state][country]['tracts'] += 1
    countryData[state][country]['pop'] += int(pop)
print('結果の書き込み')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countryData))
resultFile.close()
print('finish')
