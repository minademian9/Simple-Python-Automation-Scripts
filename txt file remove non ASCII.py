import pandas as pd
import os


all_files = os.listdir("ALE")

os.chdir("ALE")


txt_files = filter(lambda x: x[-4:] == '.txt', all_files)

simple_list = []
for t in txt_files:
    thefile = open(t)
    num_lines = sum(1 for line in thefile)
    thefile = open(t)
    content = thefile.readlines()
    for i in range(num_lines):
        simple_list.append(content[i].split('|'))

    # --- To Remove non ASCII Characters ------
    # for l in simple_list:
    #     for item in l:
    #         item = ''.join([c if ord(c) < 128 else '.. ' for c in item])

    df=pd.DataFrame(simple_list)

    ew = pd.ExcelWriter(t[:-4]+".xlsx",options={'encoding':'utf-8'},engine="openpyxl")
    df.to_excel(ew,index=False,header=False)
    ew.save()

    del ew
    del df
#     df.to_excel(t[:-4]+".xlsx",index=False,header=False)
