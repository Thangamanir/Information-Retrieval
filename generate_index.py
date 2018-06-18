## Program to generate index from the given data sets
##
## Author: R.Thangamani
## Student ID : 2016HT13668


import os
import re
import ntpath
import glob
from nltk.corpus import stopwords
from docx import Document
frequency = {}
occurences = []
existing_index = {}
def find_occurences(word,para_full):
    para_edited=para_full
    occur=[]
    position=0
    while (para_edited.find(word,position+len(word)+1)>=0):
        if(position!=0):
            position = para_edited.find(word,position)
        else:
            position = para_edited.find(word)
        occur.append(position)
        position = position+len(word)+1
    return occur

def remove_stopwords(word):
    word=word.lower()
    if word not in stopwords.words("english"):
        return word
    else:
        return ""

folder_Name = os.path.dirname(os.path.abspath(__file__))
for sampleFile in os.listdir(folder_Name):
    if sampleFile.endswith(".docx"):
        f=open(sampleFile,"rb")
        document = Document(f)
        para_full=""
        for para in document.paragraphs:
            para_full+=para.text

        match_pattern = re.findall(r'\b[a-z]{3,15}\b', para_full)
        fileName = ntpath.basename(f.name)
        for word in match_pattern:
            word = remove_stopwords(word)
            if word!="":
                if word not in frequency:
                    occurences.append(find_occurences(word,para_full))
                    frequency[word]={fileName:occurences}
                elif fileName not in frequency[word]:
                    occurences.append(find_occurences(word,para_full))
                    try:
                        frequency[word][fileName] = occurences
                    except KeyError:
                        frequency[word]={fileName:occurences}
                occurences=[]

frequency_list = frequency.keys()

for words in sorted(frequency_list):
    print(words, frequency[words])

print("Number of Dictionary terms generated= "+str(len(frequency)))
