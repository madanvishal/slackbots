import sys
import xlrd
import nltk
from nltk.tokenize import sent_tokenize
from nltk.tokenize import word_tokenize
from nltk.probability import FreqDist
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
from nltk.stem.wordnet import WordNetLemmatizer

lem = WordNetLemmatizer()
ps = PorterStemmer()
stop_words=set(stopwords.words("english"))
print(stop_words)
workbook1 = xlrd.open_workbook(r"How did we do as a Recruiter_.xlsx", on_demand=True)
sheet1 = workbook1.sheet_by_name("Form1")

for row_num in range(1,sheet1.nrows):
    row_value = sheet1.row_values(row_num)
    #print(row_value)
    print(" ")
    print("ID :",row_value[0])
    print(" ")
    tokenized_text = sent_tokenize(row_value[9])
    tokenized_word = word_tokenize(row_value[9])
    print("Tokenized Word    :",tokenized_word)
    fdist = FreqDist(tokenized_word)
    #print(fdist)
    #print(fdist.most_common(2))
    filtered_sent=[]
    for w in tokenized_word:
        if w not in stop_words:
            filtered_sent.append(w)
    print("Tokenized Sentence:",tokenized_text)
    print("Filterd Sentence  :",filtered_sent)

    stemmed_words=[]
    for w in filtered_sent:
        stemmed_words.append(ps.stem(w))
    print("Stemmed Sentence  :",stemmed_words) 
    lemmatized_words=[]
    for w in filtered_sent:
        lemmatized_words.append(lem.lemmatize(w,"v"))
    print("Lemmatized Sentence  :",lemmatized_words)

