import os
import shutil


# countries = [
# "Algeria", "Angola", "Argentina", "Australia", "Austria",
# "Azerbaijan", "Bahrain", "Bangladesh", "Belgium", "Bermuda",
#  "Brazil", "Cambodia", "Canada", "Chile", "China", "Colombia",
#   "Congo", "Cote d'Ivoire", "Croatia", "Czech Republic", "Denmark",
#   "Dominican Republic", "Ecuador", "Egypt", "Equatorial Guinea", "Estonia",
#   "Ethiopia", "Finland", "France", "Germany", "Ghana", "Greece", "Hong Kong",
#    "Hungary", "India", "Indonesia", "Iraq", "Ireland", "Israel", "Italy", "Japan",
#    "Jordan", "Kazakhstan", "Kenya", "Korea, Republic of", "Kuwait",
#    "Lao People's Democratic Republic", "Lebanon", "Libya", "Lithuania",
#    "Luxembourg", "Malaysia", "Mauritius", "Mexico", "Mongolia", "Morocco",
#     "Mozambique", "Myanmar", "Netherlands", "New Zealand", "Nigeria", "Norway",
#      "Oman", "Pakistan", "Panama", "Papua New Guinea",
#    "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania",
#     "Russian Federation", "Saudi Arabia", "Serbia", "Singapore",
#     "Slovakia", "South Africa", "Spain", "Sweden", "Switzerland",
#     "Taiwan Province of China", "Thailand", "Tunisia", "Turkey",
#     "Turkmenistan", "Ukraine",
#      "United Arab Emirates", "United Kingdom",
#      "Uruguay", "Venezuela Bolivarian Republic of",
#      "Vietnam", "Yemen"
# ]

countries = [
"Algeria", "Angola", "Argentina", "Australia", "Austria",
"Azerbaijan", "Bahrain", "Bangladesh", "Belgium", "Bermuda",
 "Brazil", "Cambodia", "Canada", "Chile", "China", "Colombia",
  "Congo", "Cote d'Ivoire", "Croatia", "Czech Republic", "Denmark",
  "Dominican Republic", "Ecuador", "Egypt", "Equatorial Guinea", "Estonia",
  "Ethiopia", "Finland", "France", "Germany", "Ghana", "Greece", "Hong Kong",
   "Hungary", "India", "Indonesia", "Iraq", "Ireland", "Israel Sprint", "Italy", "Japan",
   "Jordan", "Kazakhstan", "Kenya", "Korea", "Kuwait",
   "Laos", "Lebanon", "Libya", "Lithuania",
   "Luxembourg", "Malaysia", "Mauritius", "Mexico", "Mongolia", "Morocco",
    "Mozambique", "Myanmar", "Netherlands", "New Zealand", "Nigeria", "Norway",
     "Oman", "Pakistan", "Panama", "Papau New Guinea",
   "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania",
    "Russia", "Saudi Arabia", "Serbia", "Singapore",
    "Slovakia", "South Africa", "Spain", "Sweden", "Switzerland", "Taiwan", "Thailand", "Tunisia", "Turkey", "Turkmenistan", "Ukraine",
     "UAE", "UK", "Uruguay", "Venezuela", "Vietnam", "Yemen","USA","Puerto Rico"
]

files_regions = [ "SSA", "APAC", "ASEAN","Canada",
"Europe","Greater China",
"India", "LatAm", "MENAT",
"RUCIS","US"]

regions = [ "SSA", "APAC", "ASEAN","Canada",
"Europe","Greater China",
"India region", "LATAM", "MENAT",
"RUCIS","US"]

path = os.getcwd()
path_child = path + "\\rename"
files = os.listdir(path_child)



os.chdir(path_child)

# Sprint Country Template_ China.xlsm

for c in countries:
    src_dir=path_child+"\\Template.xlsm"
    dst_dir=path_child+"\\Sprint Country Template_ "+str(c)+".xlsm"
    shutil.copy(src_dir,dst_dir)

for r in regions:
    src_dir=path_child+"\\Template.xlsm"
    dst_dir=path_child+"\\Sprint Country Template_ "+str(r)+".xlsm"
    shutil.copy(src_dir,dst_dir)

'''


regions = [ "Africa", "ANZ", "ASEAN", "Canada",
"Europe", "Greater China",
"India", "Latin America", "MENAT",
"Russia/CIS",
"United States"]

'''
