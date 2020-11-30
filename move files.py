import os, os.path
import shutil

path = os.getcwd()
path_child = path + "\\CONFIDENTIAL - Project Sprint Country Files"
os.chdir(path_child)


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

for c in countries:
    src_dir=path+"\\Rename"+"\\Sprint Country Template_ "+str(c)+".xlsm"
    # dst_dir=path_child+"\\"+str(c)+box_path+"\\"+str(c)+".xlsm"
    dst_dir=path_child+"\\All_Countries\\"+str(c)+"\\Sprint Country Template_ "+str(c)+".xlsm"
    shutil.copy(src_dir,dst_dir)

for i  in range(len(files_regions)):
    src_dir=path+"\\Rename"+"\\Sprint Country Template_ "+str(regions[i])+".xlsm"
    # dst_dir=path_child+"\\"+str(c)+box_path+"\\"+str(c)+".xlsm"
    dst_dir=path_child+"\\Sprint_Regions"+"\\"+str(files_regions[i])+"\\Sprint Country Template_ "+str(regions[i])+".xlsm"
    shutil.copy(src_dir,dst_dir)
