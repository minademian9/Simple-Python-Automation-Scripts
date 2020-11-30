import os
import shutil

path = os.getcwd()
all_files = os.listdir(path)

xlsb_files = filter(lambda x: x[-5:] == '.xlsb', all_files)


for f in xlsb_files:
    os.rename(f,f.replace(" ", "_"))
