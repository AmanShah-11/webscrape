from argparse import ArgumentParser
import argparse
from bs4 import BeautifulSoup
from collections import Counter
import csv
import datetime as dt
from datetime import datetime
import gensim
from gensim.utils import simple_preprocess
from gensim.models import CoherenceModel
from gensim.models.coherencemodel import CoherenceModel
from gensim import corpora as corpora
from gensim.models import LsiModel
import json
import logging
import logging.config
from itertools import chain
import matplotlib.pyplot as plt
import matplotlib
import nltk
from nltk import word_tokenize
from nltk.tokenize import RegexpTokenizer
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
import numpy as np
import os.path
import os
import pandas as pd
from PIL import Image
import PIL
from pprint import pprint
import pyLDAvis
import pyLDAvis.gensim
import re
import requests
from requests import get
from schema import SCHEMA
import selenium
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium import webdriver as wd
import shutil
import smtplib
import string
from string import Template
import wordcloud
import time
import unittest
from urllib.request import urlopen
import urllib
import win32com.client
from wordcloud import WordCloud, ImageColorGenerator
from wordcloud import STOPWORDS
import xlsxwriter

punctuations = '!()-[]{};:"\,<>./?@#$%^&*_~'

# import pywin32

x = datetime.today()


url_indeed = []
url_indeed.append("https://ca.indeed.com/cmp/Caa/reviews")
url_indeed.append("https://ca.indeed.com/cmp/Caa/reviews?start=20")
url_indeed.append("https://ca.indeed.com/cmp/Caa/reviews?start=40")
url_glassdoor = "https://www.glassdoor.ca/Reviews/CAA-South-Central-Ontario-Reviews-E150598.htm"

stars = []
title = []
description = []
date = []
position = []
location = []
pros = []
cons = []

number_of_topics = 5
words = 10

start = time.time()

DEFAULT_URL = ('https://www.glassdoor.ca/Reviews/CAA-South-Central-Ontario-Reviews-E150598.htm')


parser = ArgumentParser()
parser.add_argument('-u', '--url',
                        help='URL of the company\'s Glassdoor landing page.',
                        default=DEFAULT_URL)
parser.add_argument('-f', '--file', default='glassdoor_ratings.csv',
                        help='Output file.')
parser.add_argument('--headless', action='store_true',
                    help='Run Chrome in headless mode.')
parser.add_argument('--username', help='Email address used to sign in to GD.')
parser.add_argument('-p', '--password', help='Password to sign in to GD.')
parser.add_argument('-c', '--credentials', help='Credentials file')
parser.add_argument('-l', '--limit', default=152,
                        action='store', type=int, help='Max reviews to scrape')
parser.add_argument('--start_from_url', action='store_true',
                        help='Start scraping from the passed URL.')
parser.add_argument(
    '--max_date', help='Latest review date to scrape.\
    Only use this option with --start_from_url.\
    You also must have sorted Glassdoor reviews ASCENDING by date.',
    type=lambda s: dt.datetime.strptime(s, "%Y-%m-%d"))
parser.add_argument(
    '--min_date', help='Earliest review date to scrape.\
    Only use this option with --start_from_url.\
    You also must have sorted Glassdoor reviews DESCENDING by date.',
    type=lambda s: dt.datetime.strptime(s, "%Y-%m-%d"))
args = parser.parse_args()

if not args.start_from_url and (args.max_date or args.min_date):
    raise Exception(
       'Invalid argument combination:\
        No starting url passed, but max/min date specified.'
    )
elif args.max_date and args.min_date:
    raise Exception(
          'Invalid argument combination:\
           Both min_date and max_date specified.'
     )

if args.credentials:
    with open(args.credentials) as f:
        d = json.loads(f.read())
        args.username = d['username']
        args.password = d['password']
else:
    try:
        with open('secrets.txt') as f:
            d = json.loads(f.read())
            args.username = d['username']
            args.password = d['password']
    except FileNotFoundError:
        msg = 'Please provide Glassdoor credentials.\
        Credentials can be provided as a secret.txt file in the working\
        directory, or passed at the command line using the --username and\
        --password flags.'
        raise Exception(msg)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
logger.addHandler(ch)
formatter = logging.Formatter(
    '%(asctime)s %(levelname)s %(lineno)d\
    :%(filename)s(%(process)d) - %(message)s')
ch.setFormatter(formatter)

logging.getLogger('selenium').setLevel(logging.CRITICAL)
logging.getLogger('selenium').setLevel(logging.CRITICAL)

myy_dictionary = {
    "1": 'January',
    "2": 'February',
    "3": 'March',
    "4": 'April',
    "5": 'May',
    "6": 'June',
    "7": 'July',
    "8": 'August',
    "9": 'September',
    "10": 'October',
    "11": 'November',
    "12": 'December'
}

def make_directory():
    # # dir = os.path.join("C:\\Users\\asha1\\WebScrape\\" + str(myy_dictionary[str(x.month)]) + " " + str(x.day) + " " + str(x.year))
    # # if not os.path.exists(dir):
    # #     os.mkdir(dir)
    # Creates a directory if needed for a new month
    dir_month_move = os.path.join("C:\\Users\\asha1\\Previous\\" + str(myy_dictionary[str(x.month)]) + " " + str(x.year))
    if not os.path.exists(dir_month_move):
        os.mkdir(dir_month_move)

    #Creates a new directory if needed for a new day
    dir_move = os.path.join(dir_month_move + "\\" + str(myy_dictionary[str(x.month)]) + " " + str(x.day) + " " + str(x.year))
    if not os.path.exists(dir_move):
        os.mkdir(dir_move)
    return dir_move

# Moves old files from the same day to "previous" location
def old_file_move(dst):
    files = []
    filespath = []
    countloop = int(0)
    path = os.path.join("C:\\Users\\asha1\\WebScrape")

    for r, d, f in os.walk(path):
        for file in f:
            if 'glassdoor' in file or 'indeed' in file or "coherence" in file:
                files.append(file)
                filespath.append(r)

    for f in files:
        src = str(filespath[countloop])
        shutil.move(os.path.join(src,f), os.path.join(dst,f))
        countloop = countloop + 1

# Webscrapes CAA indeed pages
def web_scrape_indeed():
    for page in range(0, 3):
        print("New Page")
        response = get(url_indeed[page])
        html_soup = BeautifulSoup(response.text, 'html.parser')
        # print(html_soup.prettify())

        # Finds all the elements on the page
        star_containers = html_soup.find_all('div', class_='cmp-ReviewRating-text')
        review_containers = html_soup.find_all('div', class_='cmp-Review-text')
        title_containers = html_soup.find_all('div', class_='cmp-Review-title')
        # pros_containers = html_soup.find_all('div', class_='cmp-ReviewProsCons-prosText')
        # cons_containers = html_soup.find_all('div', class_='cmp-ReviewProsCons-consText')
        position_containers = html_soup.find_all('span', class_='cmp-ReviewAuthor')
        general_review = html_soup.find_all('div', class_='cmp-Review-content')
        # location_containers = html_soup.find_all('span', class_='cmp-ReviewAuthor')
        # date_containers = html_soup.find_all('span', class_='cmp-ReviewAuthor')

        #Test to see if the containers are getting the proper amount of reviews
        print(len(star_containers))
        print(len(review_containers))
        print(len(title_containers))
        print(len(position_containers))
        print(len(general_review))
        # print(len(cons_containers))
        # print(len(location_containers))
        # print(len(date_containers))

        # Adds each element on the page to list in order
        for i in range(0, len(title_containers) - 1):
            first_star = star_containers[i].text
            stars.append(first_star)

            first_review_text = review_containers[i].span.span.text
            description.append(first_review_text)

            first_title_text = title_containers[i].text
            title.append(first_title_text)

            first_position_text = position_containers[i].text
            position.append(first_position_text)

            try:
                pros_containers = general_review[i].find('div', class_='cmp-ReviewProsCons-prosText')
                first_pros_text = pros_containers.span.text
                pros.append(first_pros_text)
            except Exception as e:
                pros.append("N/A")

            try:
                cons_containers = general_review[i].find('div', class_='cmp-ReviewProsCons-consText')
                first_cons_text = cons_containers.span.text
                cons.append(first_cons_text)
            except Exception as e:
                cons.append("N/A")
            # first_location_text = location_containers[i].a.text.
            # location.append(first_location_text)
            # #
            # first_date_text = date_containers[i].a.text
            # date.append(first_date_text)
    # Creates dataframe for all given elements and exports to excel
    test_df = pd.DataFrame({'Stars': stars, 'Title': title, 'Description': description, 'Position, Date, Location': position, "Pros": pros, "Cons": cons})
                             # 'Date': date, 'Location': location})
    # test_df.head(10)
    # print(test_df.head(10))

    # test_df.loc[:, 'Stars'] = test_df['Stars'].str[0:3]

    export_csv = test_df.to_csv('C:\\Users\\asha1\\WebScrape\\indeed.csv')


# Emails to HRC using VBA Script in Excel Document
def email_attachments():
    if os.path.exists("C:\\Users\\asha1\\WebScrape\\Email.xlsm"):
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Workbooks.Open(os.path.abspath("C:\\Users\\asha1\\WebScrape\\Email.xlsm"), ReadOnly=1)
        xl.Application.Run("Send_the_Email")
        del xl
    else:
        print("Path doesn't exist")


# Creates csv file with only the text from the glassdoor webscraping
def csv_convert(path, file_name):
    df = pd.read_csv(path + file_name)
    saved_column_pros = df.pros
    saved_column_cons = df.cons
    saved_column_MainText = df.MainText
    res = pd.DataFrame([], [])
    res = res.append(saved_column_pros)
    res = res.append(saved_column_cons)
    res = res.append(saved_column_MainText)
    print(Counter(" ".join(df['pros']).split()).most_common(100))
    res.to_csv("C:\\Users\\asha1\\WebScrape\\glassdoortext"  + ".csv")
    print(res)
    return 0


# Outputs the most common words from glassdoor webscraping
# Not used anymore because run time is too long and doesn't exlude punctuation
def most_common_glassdoor():
    top_N = 60
    df = pd.read_csv('C:\\Users\\asha1\\WebScrape\\glassdoor'  + str(x.year) + str(x.month) + str(x.day) + '.csv',
                     usecols=['pros', 'cons', 'MainText'])

    txt = df.pros.str.lower().str.replace(r'\|', ' ').str.cat(sep=' ') + df.cons.str.lower().str.replace(r'\|', ' ').str.cat(sep=' ') + df.MainText.str.lower().str.replace(r'\|', ' ').str.cat(sep=' ')
    words = nltk.tokenize.word_tokenize(txt)
    word_dist = nltk.FreqDist(words)

    stopwords = nltk.corpus.stopwords.words('english')
    words_except_stop_dist = nltk.FreqDist(w for w in words if w not in stopwords)

    rslt = pd.DataFrame(words_except_stop_dist.most_common(top_N),
                        columns=['Word', 'Frequency']).set_index('Word')
    print(rslt)
    rslt.to_csv('C:\\Users\\asha1\\WebScrape\\glassdoorcommon'  + str(x.year) + str(x.month) + str(x.day) + '.csv')

# Outputs the most common words from glassdoor webscraping
def most_common_glassdoor_text(texts):
    # text_counter = collections.Counter(texts)
    #     # Common = text_counter.most_common(300)
    Common = Counter(chain.from_iterable(texts)).most_common(60)
    print(Common)
    df = pd.DataFrame(Common, columns=['Common Words in Glassdoor','Amount of Occurences'])
    df.to_csv('C:\\Users\\asha1\\WebScrape\\glassdoorcommonwords'  +  '.csv')



# Outputs the most common words from indeed scraping
# NOt used anymore because it takes too long and doesn't exclude punctuation
def most_common_indeed():
    top_N = 60
    df = pd.read_csv('C:\\Users\\asha1\\WebScrape\\indeed'  + str(x.year) + str(x.month) + str(x.day) + '.csv',
                     usecols=['Description', 'Title'])

    txt = df.Description.str.lower().str.replace(r'\|', ' ').str.cat(sep=' ') + df.Title.str.lower().str.replace(r'\|', ' ').str.cat(sep=' ')
    words = nltk.tokenize.word_tokenize(txt)
    word_dist = nltk.FreqDist(words)

    stopwords = nltk.corpus.stopwords.words('english')
    words_except_stop_dist = nltk.FreqDist(w for w in words if w not in stopwords)

    rslt = pd.DataFrame(words_except_stop_dist.most_common(top_N),
                        columns=['Word', 'Frequency']).set_index('Word')
    print(rslt)
    rslt.to_csv('C:\\Users\\asha1\\WebScrape\\indeedcommon'  + str(x.year) + str(x.month) + str(x.day) + '.csv')

# Outputs the most common words from indeed scraping
def most_common_indeed_text(texts):
    Common = Counter(chain.from_iterable(texts)).most_common(60)
    print(Common)
    df = pd.DataFrame(Common, columns=['Common Words in Indeed','Number of Occurences'])
    df.to_csv('C:\\Users\\asha1\\WebScrape\\indeedcommonwords'  +  '.csv')

# Creates a wordcloud for the indeed text
def most_common_wordcloud_indeed(texts):
    text = ""
    for i in texts:
        print (i)
        text = text + str(i)
    wordcloud = WordCloud(max_font_size=100, max_words=100, background_color="black").generate(text)
    plt.figure()
    plt.imshow(wordcloud, interpolation= "bilinear")
    plt.axis("off")
    # plt.show()
    # wordcloud.to_file("C:\\Users\\asha1\\Webscrape\\reviewforindeed.png")
    plt.savefig("C:\\Users\\asha1\\Webscrape\\reviewforindeed.png")
    CAA_logo = np.array(Image.open("C:\\Users\\asha1\\WebScrape\\CAA_logo.png"))
    # CAA_logo = CAA_logo.reshape((CAA_logo.shape[0], CAA_logo.shape[1]), order='F')
    transformed_CAA_logo = np.ndarray((CAA_logo.shape[0],CAA_logo.shape[1]), np.int32)

    for i in range(len(CAA_logo)):
        transformed_CAA_logo[i] = list(map(transform_format, CAA_logo[i]))
    print(transformed_CAA_logo)
    wc = WordCloud(background_color="black", max_words=500, mask=CAA_logo, contour_width=3, contour_color="blue")
    wc.generate(text)
    wc.to_file("C:\\Users\\asha1\\WebScrape\\CAA_transformed_indeed.png")
    plt.figure(figsize=[20,10])
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")
    # plt.show()
    return 0

# Creates a wordcloud for the glassdoor text
def most_common_wordcloud_glassdoor(texts):
    # Extracts all the text information into a string variable
    text = ""
    for i in texts:
        print(i)
        text = text + str(i)
    # Creates the wordcloud
    wordcloud = WordCloud(max_font_size=100, max_words=100, background_color="white").generate(text)
    plt.figure()
    plt.imshow(wordcloud, interpolation="bilinear")
    plt.axis("off")
    # plt.show()
    # wordcloud.to_file("C:\\Users\\asha1\\Webscrape\\reviewforindeed.png")
    # Saves the wordcloud
    plt.savefig("C:\\Users\\asha1\\Webscrape\\reviewforglassdoor.png")

    # Creates a logo in the shape of a CAA logo
    # Opens the CAA logo and takes its shape
    CAA_logo = np.array(Image.open("C:\\Users\\asha1\\WebScrape\\CAA_logo.png"))
    transformed_CAA_logo = np.ndarray((CAA_logo.shape[0], CAA_logo.shape[1]), np.int32)

    # Changes the array values of the array to the value 255
    for i in range(len(CAA_logo)):
        transformed_CAA_logo[i] = list(map(transform_format, CAA_logo[i]))
    print(transformed_CAA_logo)

    # Creates the wordclud in the CAA logo
    wc = WordCloud(background_color="black", max_words=500, mask=CAA_logo, contour_width=3, contour_color="blue")
    wc.generate(text)
    # Saves it to the file
    wc.to_file("C:\\Users\\asha1\\WebScrape\\CAA_transformed_glassdoor.png")
    # Creates the dimensions
    plt.figure(figsize=[20, 10])
    plt.imshow(wc, interpolation="bilinear")
    # Makes sure that it has no axises
    plt.axis("off")
    return 0

# Transforms the values of the numpy array of the CAA logo from 0 to 255 to create the wordcloud
def transform_format(val):
    return int(255)

# All the functions  for the LSA analysis on the webscraped pages
# Reads all the text from the indeed excel document
def load_data(path, file_name):
    print("load_data")
    documents_list = []
    titles = []
    with open(os.path.join(path, file_name), "r", encoding='utf8', errors='ignore') as fin:
        for line in fin.readlines():
            text = line.strip()
            documents_list.append(text)
    print(len(documents_list))
    titles.append(text[0:min(len(text), 250)])
    return documents_list, titles

# Reads all the text from the glassdoor excel document
def load_data_glassdoor(path, file_name):
    print("load data")
    documents_list = []
    titles = []
    with open(os.path.join(path, file_name), "r", encoding='utf8', errors='ignore') as fin:
        for line in fin.readlines():
            text = line.strip()
            documents_list.append(text)
    print(len(documents_list))
    titles.append(text[0:min(len(text), 250)])
    return documents_list, titles

# Gets rid of stop words, punctuation and makes text lower case
def preprocess_data(doc_set):
    punctuations = '!()-[]{};:"\,<>./?@#$%^&*_~'
    print("preprocess data")
    tokenizer = RegexpTokenizer(r'\w+')
    # create English stop words list
    en_stop = set(stopwords.words('english'))
    en_punctuation = set(string.punctuation)
    # Create p_stemmer of class PorterStemmer
    p_stemmer = PorterStemmer()
    # list for tokenized documents in loop
    texts = []
    # loop through document list
    for i in doc_set:
        # clean and tokenize document string
        raw = i.lower()
        for x in raw:
            if x in punctuations:
                raw = raw.replace(x,"")
        tokens = tokenizer.tokenize(raw)
        # remove stop words from tokens
        stopped_tokens = [i for i in tokens if i not in en_stop]
        # stem tokens
        stemmed_tokens = [p_stemmer.stem(i) for i in stopped_tokens]
        # add tokens to list
        texts.append(stemmed_tokens)
    return texts

def words(filepath):
    punctuations = '!()-[]{};:"\,<>./?@#$%^&*_~'
    tokenizer = RegexpTokenizer(r'\w+')
    en_stop = set(stopwords.words('english'))
    en_punctuation = set(string.punctuation)
    p_stemmer = PorterStemmer()
    texts = []
    word_list = []
    try:
        with open(filepath, "r", encoding='utf8', errors='ignore') as fin:
            for line in fin.readlines():
                line = line.replace("\n", " ")
                word_list.append(line)
    except Exception as e:
        print(e)
    for i in word_list:
        # clean and tokenize document string
        raw = i.lower()
        for x in raw:
            if x in punctuations:
                raw = raw.replace(x,"")
        tokens = tokenizer.tokenize(raw)
        # remove stop words from tokens
        stopped_tokens = [i for i in tokens if i not in en_stop]
        # stem tokens
        stemmed_tokens = [p_stemmer.stem(i) for i in stopped_tokens]
        # add tokens to list
        texts.append(stemmed_tokens)
    for text in texts:
        for text_indiv in text:
            print(text_indiv + "TEXT_INDIVIDUAL")
    return texts

def word_list_compare(word_list, doc_clean):
    doc_clean_emotion = []
    print("does it work")
    for i in doc_clean:
        print(i)
        for x in i:
            print(x)
            for array in word_list:
                for word in array:
                    if x == word:
                        doc_clean_emotion.append(x)
    return doc_clean_emotion

def word_common(doc_clean_emotion, save_file_as):
    Common = Counter(doc_clean_emotion).most_common(30)
    df = pd.DataFrame(Common, columns = ["Most common words", "Number of Occurences of Words"])
    df.to_csv("C:\\Users\\asha1\\WebScrape\\" + save_file_as, index=False)

# Creates a csv file with all the emotional words(happy or sad) in it
def emotional_words(filepath_emotion,  doc_clean, save_file_as):
    #Obtains all the words from the emotinal txt file
    word_txt = words(filepath_emotion)
    # Compares the words obtained through webscraping and the emotiona txt file
    common_list = word_list_compare(word_txt, doc_clean)
    # puts the most common words into a dataframe that's exported to a csv file
    word_common(common_list, save_file_as)

# Creates matrix of how often words occur and dictionary of all the words that exist
def prepare_corpus(doc_clean):
    print("print corpus")
    """
      Input  : clean document
      Purpose: create term dictionary of our courpus and Converting list of documents (corpus) into Document Term Matrix
      Output : term dictionary and Document Term Matrix
      """
    # Creating the term dictionary of our courpus, where every unique term is assigned an index.
    dictionary = corpora.Dictionary(doc_clean)
    # Converting list of documents (corpus) into Document Term Matrix using dictionary prepared above.
    doc_term_matrix = [dictionary.doc2bow(doc) for doc in doc_clean]
    # generate LDA model
    return dictionary, doc_term_matrix

# Machine learning model that creates topic modelling from the document
def create_gensim_lsa_model(doc_clean, number_of_topics, words):
    """
        Input  : clean document, number of topics and number of words associated with each topic
        Purpose: create LSA model using gensim
        Output : return LSA model
        """
    dictionary, doc_term_matrix = prepare_corpus(doc_clean)
    # generate LSA model
    lsamodel = LsiModel(doc_term_matrix, num_topics=number_of_topics, id2word=dictionary)  # train model
    print(lsamodel.print_topics(num_topics=number_of_topics, num_words=words))
    return lsamodel

# Computes the value of
def compute_coherence_values(dictionary, doc_term_matrix, doc_clean,  stop, start, step):

    # Input : dictionary : Gensim dictionary
    #         corpus: gensim corpus
    #         texts: list of input texts
    #         stop: max num of topics
    # Purpse : Compute c_v coherence for different number of topics
    # Output: model_list : List of LSA topic models
    #         coherence_values : Coherence values corresponding to the lDA model with respective numbers

    coherence_values = []
    model_list = []
    for num_topics in range(start, stop, step):
        model = LsiModel(doc_term_matrix, num_topics=number_of_topics, id2word=dictionary)
        model_list.append(model)
        coherencemodel = CoherenceModel(model=model, texts=doc_clean, dictionary=dictionary, coherence='c_v')
        coherence_values.append(coherencemodel.get_coherence())
    return model_list, coherence_values

# Creates csv file with information on the topic modelling words and coherence scores for indeed
def LSA_indeed_model_to_csv(model_list):
    # df = pd.DataFrame(data={"Topic Modelling Indeed":[model_list]})
    # df.to_csv("C:\\Users\\asha1\\Webscrape\\TopicModelindeed"  +  ".csv", sep=" ", index=False)
    df = pd.DataFrame.from_records(model_list)
    df.columns = ["Topic Modelling Indeed", "col 2"]
    df.applymap(str)
    df['Topic Modelling Indeed'] = df['Topic Modelling Indeed'].astype(str)
    # df['Topic Modelling Indeed'].apply(str)
    # Removes coherence scores from the dataframe
    print(df.dtypes)
    df['Topic Modelling Indeed'] = df['Topic Modelling Indeed'].str.replace('\d+', ' ')
    df['col 2'] = df['col 2'].str.replace('\d+', ' ')
    df['col 2'] = df['col 2'].str.replace(r'[^\w\s]+', '')
    # df.transpose()
    df.to_csv("C:\\Users\\asha1\\Webscrape\\TopicModelindeed"  +  ".csv", sep=" ", index=False)

# Creates csv file with information on the topic modelling words and coherence scores for indeed
def LSA_glassdoor_model_to_csv(model_list):
    df = pd.DataFrame.from_records(model_list)
    df.columns = ["Topic Modelling Glassdoor", "col 2"]
    df.applymap(str)
    df['Topic Modelling Glassdoor'] = df['Topic Modelling Glassdoor'].astype(str)
    # Removes coherence scores from the dataframe
    df['Topic Modelling Glassdoor'] = df['Topic Modelling Glassdoor'].str.replace('\d+', ' ')
    df['col 2'] = df['col 2'].str.replace('\d+', ' ')
    # Removes the punctuation from the dataframe
    df['col 2'] = df['col 2'].str.replace(r'[^\w\s]+', '')
    df.to_csv("C:\\Users\\asha1\\Webscrape\\TopicModelglassdoor"  +".csv", sep=" ", index=False)

# Creates LDA model specifically for indeed
def LDA_model_indeed(dictionaryy, doc_term_matrixx, doc_clean):
    exclude = '!()-[]{};:"\,<>./?@#$%^&*_~+*'
    lda_model = gensim.models.ldamodel.LdaModel(corpus=doc_term_matrixx,
                                                id2word=dictionaryy, per_word_topics=True, num_topics = 5)
                                                # ,
                                                # num_topics=6,
                                                # random_state=100,
                                                # update_every=1,
                                                # chunksize=100,
                                                # passes=10,
                                                # alpha='auto',
                                                # per_word_topics=True)
    pprint(lda_model.print_topics())
    doc_lda = lda_model[doc_term_matrixx]
    coherence_model_lda = CoherenceModel(model=lda_model, texts=doc_clean, dictionary=dictionaryy, coherence='c_v')
    coherence_lda = coherence_model_lda.get_coherence()
    print('\nCoherence Score: ', coherence_lda)
    raw = lda_model.print_topics()
    # try:
    #     raw = ''.join(ch for ch in x if ch not in exclude)
    # except:
    #     pass
    df = pd.DataFrame(raw)
    df.columns = ["Topic Nummber", "Topic Words"]
    df.applymap(str)
    df['Topic Words'] = df['Topic Words'].astype(str)
    # Gets rid of all the numbers
    df['Topic Words'] = df['Topic Words'].str.replace('\d+', '')
    # Gets rid of all the punctuation from the topic model
    df['Topic Words'] = df['Topic Words'].str.replace(r'[^\w\s]+', '')
    print(df.head(10))
    df.to_csv("C:\\Users\\asha1\\Webscrape\\coherencescoresindeed" + ".csv", sep=" ", index=False)
    # mallet_path = 'C:\\Users\\asha1\\AppData\\Local\\Temp\\mallet-2.0.8.zip'  # update this path
    # ldamallet = gensim.models.wrappers.LdaMallet(mallet_path, corpus=doc_term_matrixx, num_topics=20, id2word=dictionaryy)
    #
    # pprint(ldamallet.show_topics(formatted=False))
    #
    # coherence_model_ldamallet = CoherenceModel(model=ldamallet, texts=doc_clean, dictionary=dictionaryy,
    #                                            coherence='c_v')
    # coherence_ldamallet = coherence_model_ldamallet.get_coherence()
    # print('\nCoherence Score: ', coherence_ldamallet)

# Creates LDA model specifically for glassdoor
def LDA_model_glassdoor(dictionaryy, doc_term_matrixx, doc_clean):
    # Generates the lda model
    lda_model = gensim.models.ldamodel.LdaModel(corpus=doc_term_matrixx,
                                                id2word=dictionaryy, per_word_topics=True, num_topics = 5)
                                                # ,
                                                # num_topics=6,
                                                # random_state=100,
                                                # update_every=1,
                                                # chunksize=100,
                                                # passes=10,
                                                # alpha='auto',
                                                # per_word_topics=True)
    pprint(lda_model.print_topics())
    doc_lda = lda_model[doc_term_matrixx]
    coherence_model_lda = CoherenceModel(model=lda_model, texts=doc_clean, dictionary=dictionaryy, coherence='c_v')
    coherence_lda = coherence_model_lda.get_coherence()
    print('Coherence Score:', coherence_lda)
    df = pd.DataFrame(lda_model.print_topics())
    # df.columns = ["Topic Modelling Indeed", " "]
    # df['Coherence Score'] = df['Coherence Score'].str.replace('\dt+', '')
    # df['Topic Words'] = df['Topic Words'].replace('\d+', '')
    df.columns = ["Topic Nummber", "Topic Words"]
    df.applymap(str)
    # Converts the entire dataframe to an object so it can be interpreted as a string
    df['Topic Words'] = df['Topic Words'].astype(str)
    # Gets rid of all the numbers within the string
    df['Topic Words'] = df['Topic Words'].str.replace('\d+', '')
    # Gets rid of all the punctuation within the dataframe (easier for user to read)
    df['Topic Words'] = df['Topic Words'].str.replace(r'[^\w\s]+', '')
    # Prints first 10 topics generated by the model
    print(df.head(10))
    # Makes the dataframe go to a csv file
    df.to_csv("C:\\Users\\asha1\\Webscrape\\coherencescoresglassdoor" + ".csv", sep=" ", index=False)
    # mallet_path = 'C:\\Users\\asha1\\AppData\\Local\\Temp\\mallet-2.0.8.zip'  # update this path
    # ldamallet = gensim.models.wrappers.LdaMallet(mallet_path, corpus=doc_term_matrixx, num_topics=20, id2word=dictionaryy)
    #
    # pprint(ldamallet.show_topics(formatted=False))
    #
    # coherence_model_ldamallet = CoherenceModel(model=ldamallet, texts=doc_clean, dictionary=dictionaryy,
    #                                            coherence='c_v')
    # coherence_ldamallet = coherence_model_ldamallet.get_coherence()
    # print('\nCoherence Score: ', coherence_ldamallet)

# Plots graph of coherence scores and recommends how many topics to use for
# Should not be used as helper function to determine amount of topics that should be used
def plot_graph(doc_clean, start, stop, step):
    dictionary, doc_term_matrix = prepare_corpus(doc_clean)
    model_list, coherence_values = compute_coherence_values(dictionary, doc_term_matrix, doc_clean, stop, start, step)

#    Show graph
    x = range(start, stop, step)
    plt.plot(x, coherence_values)
    plt.xlabel("Number of Topics")
    plt.ylabel("Coherence score")
    plt.legend("coherence_values", loc='best')
    plt.show()

# Webscrape functions for glassdoor
def scrape(field, review, author):

    def scrape_date(review):
        return review.find_element_by_class_name("date").text

    def scrape_emp_title(review):
        if 'Anonymous Employee' not in review.text:
            try:
                res = author.find_element_by_class_name(
                    'authorJobTitle').text.split('-')[1]
            except Exception:
                res = np.nan
        else:
            res = np.nan
        return res

    def scrape_location(review):
        try:
            res = author.find_element_by_class_name(
            'authorLocation').text
        except Exception:
            res = np.nan
        return res

    def scrape_status(review):
        try:
            res = author.text.split('-')[0]
        except Exception:
            res = np.nan
        return res

    def scrape_rev_title(review):
        try:
            res = review.find_element_by_class_name('summary').text
        except Exception:
            res = np.nan
        return res

    def scrape_years(review):
        try:
            first_par = review.find_element_by_class_name(
            'reviewBodyCell').text
            res = first_par
        except:
            print("doesn't work")
            res = np.nan
        return res

    def scrape_helpful(review):
        try:
            helpful = review.find_element_by_class_name('helpfulCount')
            res = helpful[helpful.find('(') + 1: -1]
        except Exception:
            res = 0
        return res

    def expand_show_more(section):
        try:
            more_content = section.find_element_by_class_name('moreContent')
            more_link = more_content.find_element_by_class_name('moreLink')
            more_link.click()
        except Exception:
            pass

    def scrape_pros(review):
        try:
            pros = review.find_element_by_css_selector("p.mt-0.mb-xsm.v2__EIReviewDetailsV2__bodyColor.v2__EIReviewDetailsV2__lineHeightLarge")
            expand_show_more(pros)
            res = pros.text
        except Exception:
            res = np.nan
        return res

    def scrape_cons(review):
        try:
            cons = review.find_elements_by_css_selector(
                "p.mt-0.mb-xsm.v2__EIReviewDetailsV2__bodyColor.v2__EIReviewDetailsV2__lineHeightLarge")[1]
            expand_show_more(cons)
            res = cons.text
        except Exception:
            res = np.nan
        return res

    def scrape_advice(review):
        try:
            advice = review.find_element_by_class_name('adviceMgmt')
            expand_show_more(advice)
            res = advice.text.replace('\nShow Less', '')
        except Exception:
            res = np.nan
        return res

    def scrape_overall_rating(review):
        try:
            ratings = review.find_element_by_class_name('gdStars')
            overall = ratings.find_element_by_class_name(
                'rating').find_element_by_class_name('value-title')
            res = overall.get_attribute('title')
        except Exception:
            res = np.nan
        return res

    def _scrape_subrating(i):
        try:
            ratings = review.find_element_by_class_name('gdStars')
            subratings = ratings.find_element_by_class_name(
                'subRatings').find_element_by_tag_name('ul')
            this_one = subratings.find_elements_by_tag_name('li')[i]
            res = this_one.find_element_by_class_name(
                'gdBars').get_attribute('title')
        except Exception:
            res = np.nan
        return res

    def scrape_work_life_balance(review):
        return _scrape_subrating(0)

    def scrape_culture_and_values(review):
        return _scrape_subrating(1)

    def scrape_career_opportunities(review):
        return _scrape_subrating(2)

    def scrape_comp_and_benefits(review):
        return _scrape_subrating(3)

    def scrape_senior_management(review):
        return _scrape_subrating(4)

    def scrape_maintext(review):
        try:
            maintext = review.find_element_by_class_name('mainText')
            res = maintext.text
        except Exception:
            res = np.nan
        return res

    # All the functions for scraping within a list so that they are easier to call at once
    funcs = [
        scrape_date,
        scrape_emp_title,
        scrape_location,
        scrape_status,
        scrape_rev_title,
        scrape_helpful,
        scrape_pros,
        scrape_cons,
        scrape_maintext,
        scrape_overall_rating,
        scrape_work_life_balance,
        scrape_culture_and_values,
        scrape_career_opportunities,
        scrape_comp_and_benefits,
        scrape_senior_management
    ]

    # Calls all the functions for scraping and collects into a variavle for one review
    fdict = dict((s, f) for (s, f) in zip(SCHEMA, funcs))

    return fdict[field](review)

# Calls scraping functions for glassdoor and creates dataframe
def extract_from_page():

    # Extracts all the reviews from the webpages
    def extract_review(review):
        time.sleep(2)
        author = review.find_element_by_css_selector('span.authorInfo')

        res = {}
        for field in SCHEMA:
            res[field] = scrape(field, review, author)
            time.sleep(0.1)

        print("Extracting review")
        assert set(res.keys()) == set(SCHEMA)
        return res

    # Creates the pandas dataframe, and inputs the columns from the SCHEMA.py file
    res = pd.DataFrame([], columns=SCHEMA)

    reviews = browser.find_elements_by_class_name('empReview')
    # Extracts all the reviews on the page and increases index length of array for each review scraped on page
    for review in reviews:
        data = extract_review(review)
        res.loc[idx[0]] = data
        idx[0] = idx[0] + 1
        print(idx[0])
        print("index length")

    print("Done extracting from page")
    #Arguments not passed for max and min date, but it would check and stop the process if not within bounds of dates passsed
    if args.max_date and \
        (pd.to_datetime(res['date']).max() > args.max_date) or \
            args.min_date and \
            (pd.to_datetime(res['date']).min() < args.min_date):
        date_limit_reached[0] = True

    return res

# Checks if there are more pages available to scrape
def more_pages():
    next_ = browser.find_element_by_css_selector('li.pagination__PaginationStyle__next')
    print("Found li tag")
    try:
        next_.find_element_by_tag_name('a')
        print("Element is found")
        return True
    except selenium.common.exceptions.NoSuchElementException:
        return False

#Goes to next page if next page exists
def go_to_next_page():
    next_ = browser.find_element_by_css_selector('li.pagination__PaginationStyle__next a')
    browser.get(next_.get_attribute('href'))
    page[0] = page[0] + 1

# If there reviews present, it will return that there are no reviews
def no_reviews():
    return False

# Goes to initial start page for glassdoor scraping
def navigate_to_reviews():

    browser.get(args.url)
    time.sleep(1)

    if no_reviews():
        return False

    print("Navigating to reviews")
    time.sleep(1)

    return True

# Signs into glassdoor account to prevent getting asked to sign up while scraping
def sign_in():
    # logger.info(f'Signing in to {args.username}')

    url = 'https://www.glassdoor.ca/profile/login_input.htm?userOriginHook=HEADER_SIGNIN_LINK'
    browser.get(url)

    email_field = browser.find_element_by_name('username')
    password_field = browser.find_element_by_name('password')
    submit_btn = browser.find_element_by_xpath('//button[@type="submit"]')

    email_field.send_keys(args.username)
    password_field.send_keys(args.password)
    submit_btn.click()

    time.sleep(1)

# Uses google chrome to webscrape
def get_browser():
    chrome_options = wd.ChromeOptions()
    # if args.headless:
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('log-level=3')
    browser = wd.Chrome(options=chrome_options)
    return browser

# Used to see if current page exists
def get_current_page():
    paging_control = browser.find_element_by_class_name('pagingControls')
    current = int(paging_control.find_element_by_xpath(
        '//ul//li[contains\
        (concat(\' \',normalize-space(@class),\' \'),\' current \')]\
        //span[contains(concat(\' \',\
        normalize-space(@class),\' \'),\' disabled \')]')
        .text.replace(',', ''))
    return current

# Not used since arguments for dates are not passed, but it would verify if the date are sorted and raise exceptions if not
def verify_date_sorting():
    ascending = urllib.parse.parse_qs(
        args.url)['sort.ascending'] == ['true']

    if args.min_date and ascending:
        raise Exception(
            'min_date required reviews to be sorted DESCENDING by date.')
    elif args.max_date and not ascending:
        raise Exception(
            'max_date requires reviews to be sorted ASCENDING by date.')

browser = get_browser()
page = [1]
idx = [0]
date_limit_reached = [False]

# Calls functions to webscrape glassdoor
def web_scrape_glassdoor():
    res = pd.DataFrame([], columns=SCHEMA)
    sign_in()

    if not args.start_from_url:
        reviews_exist = navigate_to_reviews()
        if not reviews_exist:
            return
    elif args.max_date or args.min_date:
        verify_date_sorting()
        browser.get(args.url)
        page[0] = get_current_page()
        print(f'Starting from page {page[0]:,}.')
        time.sleep(1)
    else:
        browser.get(args.url)
        page[0] = get_current_page()
        print(f'Starting from page {page[0]:,}.')
        time.sleep(1)
    time.sleep(4)
    reviews_df = extract_from_page()
    res = res.append(reviews_df)
    count = int(0)

    while more_pages() and len(res) < args.limit:
        # not date_limit_reached[0]:
        time.sleep(2)
        go_to_next_page()
        time.sleep(12)
        reviews_df = extract_from_page()
        res = res.append(reviews_df)
        print(len(res))

    res.to_csv('C:\\Users\\asha1\\WebScrape\\glassdoor'  + '.csv', index=False, encoding='utf-8')

# Performs latent semantic analysis for scraped pages for indeed website
def LSA_indeed():
    number_of_topics = 4
    words = 8
    text_info, title_info = load_data('C:\\Users\\asha1\\WebScrape', 'indeed' + '.csv')
    clean_text = preprocess_data(text_info)
    prep_dict, prep_matrix = prepare_corpus(clean_text)
    new_model = create_gensim_lsa_model(clean_text, 5, 10)
    new_list, num_topics = compute_coherence_values(prep_dict, prep_matrix, clean_text, 2, 10, 1)
    # plot_graph(clean_text, 2, 10, 1)
    # most_common_indeed()
    common = most_common_indeed_text(clean_text)
    most_common_wordcloud_indeed(clean_text)
    LSA_indeed_model_to_csv(new_model.print_topics(num_topics=number_of_topics, num_words=words))
    LDA_model_indeed(prep_dict, prep_matrix, clean_text)
    emotional_words("negative_words.txt", clean_text, "negativewordsindeed.csv")
    emotional_words("positive_words.txt", clean_text, "positivewordsindeed.csv")

# Performs latent semantic analysis for scraped pages on glassdoor
def LSA_glassdoor():
    number_of_topics = 5
    words = 10
    csv_convert('C:\\Users\\asha1\\WebScrape\\', 'glassdoor' + '.csv')
    text_info, title_info = load_data_glassdoor('C:\\Users\\asha1\\WebScrape', 'glassdoortext' + '.csv')
    clean_text = preprocess_data(text_info)
    prep_dict, prep_matrix = prepare_corpus(clean_text)
    new_model = create_gensim_lsa_model(clean_text, 5, 10)
    new_list, num_topics = compute_coherence_values(prep_dict, prep_matrix, clean_text, 5, 10, 1)
    most_common_wordcloud_glassdoor(clean_text)
    LSA_glassdoor_model_to_csv(new_model.print_topics(num_topics=number_of_topics, num_words=words))
    common = most_common_glassdoor_text(clean_text)
    LDA_model_glassdoor(prep_dict, prep_matrix, clean_text)
    emotional_words("negative_words.txt", clean_text, "negativewordsglassdoor.csv")
    emotional_words("positive_words.txt", clean_text, "positivewordsglassdoor.csv")

    # plot_graph(clean_text, 2, 10, 1)


# Runs the program
def run_the_program():
    # Brings old files to old directory and allows space for new files to exist in current directoryis
    dirmove = make_directory()
    # # Webscrapes indeed and glassdoor respectively and saves dataframe to csv file
    web_scrape_indeed()
    time.sleep(3)
    web_scrape_glassdoor()
    #
    # # Runs the latent semantic analysis
    LSA_indeed()
    LSA_glassdoor()

    # # Emails the files to Recruitment
    email_attachments()
    # Moves the files to the storage directory
    old_file_move(dirmove)

# #Starts the program
if __name__ == '__main__':
    run_the_program()
