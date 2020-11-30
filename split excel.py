import pandas as pd
import os

box_path = "\\Latest submission by Corp-Horizontal for country HRMs to review"
 # "<country_name>\\Latest submission by Corp-Horizontal for country HRMs to review\\"

countries = [
"Algeria", "Angola", "Argentina", "Australia", "Austria",
"Azerbaijan", "Bahrain", "Bangladesh", "Belgium", "Bermuda",
 "Brazil", "Cambodia", "Canada", "Chile", "China", "Colombia",
  "Congo", "Cote d'Ivoire", "Croatia", "Czech Republic", "Denmark",
  "Dominican Republic", "Ecuador", "Egypt", "Equatorial Guinea", "Estonia",
  "Ethiopia", "Finland", "France", "Germany", "Ghana", "Greece", "Hong Kong",
   "Hungary", "India", "Indonesia", "Iraq", "Ireland", "Israel", "Italy", "Japan",
   "Jordan", "Kazakhstan", "Kenya", "Korea, Republic of", "Kuwait",
   "Lao People's Democratic Republic", "Lebanon", "Libya", "Lithuania",
   "Luxembourg", "Malaysia", "Mauritius", "Mexico", "Mongolia", "Morocco",
    "Mozambique", "Myanmar", "Netherlands", "New Zealand", "Nigeria", "Norway",
     "Oman", "Pakistan", "Panama", "Papua New Guinea",
   "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania",
    "Russian Federation", "Saudi Arabia", "Serbia", "Singapore",
    "Slovakia", "South Africa", "Spain", "Sweden", "Switzerland", "Taiwan Province of China", "Thailand", "Tunisia", "Turkey", "Turkmenistan", "Ukraine",
     "United Arab Emirates", "United Kingdom", "Uruguay", "Venezuela Bolivarian Republic of", "Vietnam", "Yemen"
]

# regions = [ "Africa", "ANZ", "ASEAN", "Canada",
# "Europe", "Greater China",
# "India", "Latin America", "MENAT",
# "Russia/CIS",
# "United States"]

regions = [ "Africa", "ANZ", "ASEAN", "Canada",
"Europe", "Greater China",
"India", "LATAM", "MENAT",
"Russia/CIS","United States"]

files_regions = [ "SSA", "APAC", "ASEAN","Canada",
"Europe","Greater China",
"India", "LatAm", "MENAT",
"RUCIS","US"]


x = pd.ExcelFile('1.xlsm')
df = x.parse('Listing_of_all_ees')
df = df.iloc[78:]
df.columns = df.iloc[0]
df = df.iloc[1:]

df['Country'] = map(lambda z: str(z).title(), df['Country'])
# df[df['Country']=='Brazil']

# Get Current folder path
path = os.getcwd()
path_child = path + "\\CONFIDENTIAL - Project Sprint Country Files"
os.chdir(path_child)

# Country Split
for c in countries:
    country_df = df[df['Country']==c]
    if len(country_df.head()) >0:
        country_path = path_child + "\\" + c + box_path
        os.chdir(country_path)
        country_df.to_excel(c+".xlsx",index=False)
        os.chdir(path_child)

# Region Split
for i in range(len(regions)):
    regions_df = df[df['REGION_NM']==regions[i]]
    if len(regions_df.head()) >0:
        region_path = path_child + "\\_Regional Cuts\\" + files_regions[i] + "\\Latest submission by Corp-Horizontal"
        os.chdir(region_path)
        regions_df.to_excel(regions[i]+".xlsx",index=False)
        os.chdir(path_child)
