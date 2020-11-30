import shutil
import os

path = os.getcwd()
all_files = os.listdir(path)

xlsb_files = filter(lambda x: x[-5:] == '.xlsb', all_files)

for t in xlsb_files:

    try:
        os.mkdir(t[:-4])
    except Exception as e :
        print e

    print "Directory Created"

    src_dir=path+"\\"+t
    dst_dir=path+"\\"+t[:-5]

    shutil.move(src_dir,dst_dir)


    shutil.make_archive(t[:-5], 'zip', t[:-5])

    print "Zip file Created for " + t

    src_dir=path+"\\"+t[:-5]+"\\"+t
    dst_dir=path

    shutil.move(src_dir,dst_dir)

    # os.remove(path+'\\'+t[:-4])
    shutil.rmtree(path+'\\'+t[:-5])

    print "Directory Cleanup Done"
    print "."
    print "."
    print "."

raw_input("Completed...Press any key to Exit...")
