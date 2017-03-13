# A script to convert .docx files to .txt files

import textract
import os
import yaml

# Accounting for the folders that I've organized by year
num = 2007
years = []
while (num < 2018):
    years.append(str(num))
    num = num + 1

# open file with my directory information & extraction
with open("file_path.yml", 'r') as ymlfile:
    fp = yaml.load(ymlfile)

file_path = fp["mysql"]["file_path"]

# Extracting data from .docx files and saving them in .txt files within the dirs
for year in years:
    path = file_path + year + '/'
    dirs = os.listdir(path)

    for files in dirs:
        extension = os.path.splitext(files)[1]
        if  extension == '.docx' in files:
            file_name = os.path.splitext(files)[0]
            f = open(path + file_name + '.txt' , 'w')
            text = textract.process(path + files).decode("utf-8")
            f.write(text)
f.close()
