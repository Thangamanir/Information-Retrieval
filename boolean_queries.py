## Program to search for queries in an index
## User inputs the query strings via the console
## Author: R.Thangamani
## Student ID : 2016HT13668

import os
import re
import ntpath
import glob
from nltk.corpus import stopwords
from docx import Document
from functools import reduce
from defaultlist import defaultlist
from collections import defaultdict
from operator import itemgetter

frequency = {}
occurences = []


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

def search_query():
    query = input(" Enter the query : ")

    query = query.lower()
    query = query.strip()

    wordlist = query.split()
    search_words = [ x for x in wordlist if x in frequency]


    doc_has_word = [ (frequency[word].keys(),word) for word in search_words ]
    doc_words = defaultdict(list)
    for d, w in doc_has_word:
        for p in d:
            doc_words[p].append(w)



    result_set = {}

    for i in doc_words.keys():
        count = 0
        matches = len(doc_words[i])     # number of matches
        for w in doc_words[i]:
            count += len(frequency[w][i])   # count total occurances
        result_set[i] = (matches,count)



    print ("   Document \t\t Words matched \t\t Total Frequency ")
    print ('-'*40)
    for (doc, (matches, count)) in sorted(result_set.items(), key = itemgetter(1), reverse = True):
        print (doc, "\t\t",doc_words[doc],"\t\t",count)

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
search_query()
