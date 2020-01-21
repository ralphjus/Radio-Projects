import sys
import glob
import os
import re
import shutil

def get_names(path, filename):
    expression = os.path.join(path, filename+'*')
    return glob.glob(expression)

if __name__ == "__main__":

    args = sys.argv
    if len(args) != 2:
        print('specify path')
        sys.exit(-1)

    path = args[1]
    print 'Searching in {}'.format(path)

    try:
        filenames = get_names(path, "")
    except:
        print('Invalid file path')
        sys.exit(-1)

    num_regex = re.compile(r".*?([0-9]+)")

    def key_func(x):
        number = int(re.match(num_regex, x).groups()[0])
        return number

    filenames = sorted(filenames, key=key_func)

    for filename in filenames: #You are going to do the below procedure over each file
    #example filename: '/Volumes/Studios/Studio-I/SD1/Studio_I_Audio_Files/StudioI-700R_03152017184020.mp3'
        date_regex = re.compile(r".*?0([0-9]+)\.mp3")
        print 'Copying {}...'.format(filename)

        def date(x):
            number = re.match(date_regex, x).groups()[0]
            if number is None:
                print "Regex did not match"
            return int(number)
        print date(filename)

        if 4172017100000 < date(filename) < 4172017115959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ1")

        if 4172017120000 < date(filename) < 4172017135959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ2")

        if 4172017140000 < date(filename) < 4172017155959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ3")

        if 4172017160000 < date(filename) < 4172017175959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ4")

        if 4172017180000 < date(filename) < 4172017195959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ5")

        if 4172017200000 < date(filename) < 4172017215959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ6")

        if 4172017220000 < date(filename) < 4172017235959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ7")

        if 4182017100000 < date(filename) < 4182017115959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ8")

        if 4182017120000 < date(filename) < 4182017135959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ9")
        
        if 4182017160000 < date(filename) < 4182017175959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ10")

        if 4182017180000 < date(filename) < 4182017195959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ11")

        if 4182017200000 < date(filename) < 4182017235959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ12")

        if 4192017100000 < date(filename) < 4192017115959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ13")
        
        if 4192017160000 < date(filename) < 4192017175959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ14")

        if 4192017180000 < date(filename) < 4192017195959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ15")

        if 4192017200000 < date(filename) < 4192017215959:
            print 'running copy'
            shutil.copy2(filename, "Desktop/radio/recordings/DJ16")




#python ~/Downloads/sortfilenames.py /Volumes/Studios/Studio-I/SD1/Studio_I_Audio_Files
#python ~/Downloads/sortfilenames.py ~/Downloads/airchecks
