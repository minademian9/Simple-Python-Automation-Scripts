import os, os.path
import win32com.client

def run_macro(the_excel_path):
    filename = the_excel_path

    if os.path.exists(filename):
        try:
            workbook=win32com.client.Dispatch("Excel.Application")
            ws = workbook.Workbooks.Open(os.path.abspath(filename), ReadOnly=1)
            # ws = workbook.Workbooks.Open(os.path.abspath(filename))
            workbook.Application.Run("'"+filename+"'"+"!Module2.Write_sprint")
            # ws.Save()
            workbook.Application.Quit()
            del workbook
            del ws
        except Exception as e:
            print e
            print filename





    # filename = "India.xlsm"
    #
    # if os.path.exists(filename):
    #     workbook=win32com.client.Dispatch("Excel.Application")
    #     ws = workbook.Workbooks.Open(os.path.abspath(filename), ReadOnly=1)
    #     workbook.Application.Run(filename+"!Module2.Write_sprint")
    #     # workbook.Application.Save()
    #     ws.Save()
    #     # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    #     workbook.Application.Quit() # Comment this out if your excel script closes
    #     del workbook


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

counter = 0.0

for c in countries:
    country_path = path_child + "\\All_Countries\\" + c
    os.chdir(country_path)
    file_name="Sprint Country Template_ "+str(c)+".xlsm"
    # file_name=str(c)+".xlsm"
    run_macro(file_name)
    counter += 1
    print file_name + " Done ! " + str(counter/len(countries)*100) + " % Completed"
    os.chdir(path_child)

print "\n"
print "Countries Done"
print "\n"

counter = 0.0
for i  in range(len(files_regions)):
    region_path = path_child + "\\Sprint_Regions\\" + files_regions[i]
    os.chdir(region_path)
    file_name="Sprint Country Template_ "+str(regions[i])+".xlsm"
    # file_name=str(c)+".xlsm"
    run_macro(file_name)
    counter += 1
    print file_name + " Done ! " + str(counter/len(files_regions)*100) + " % Completed"
    os.chdir(path_child)

print "\n"
print "Regions Done"
print "\n"

print "Quitting..."
# workbook.Application.Quit()
# del workbook
