import pandas as pd
import requests
from bs4 import BeautifulSoup
from nltk.tokenize import word_tokenize, sent_tokenize
import re
import os
import nltk

nltk.download('punkt')

e_data = pd.read_excel('Input.xlsx')
urls = [x for x in e_data.iloc[:, 1]]

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
}


for i, url in enumerate(urls):
    r = requests.get(url, headers=headers).text
    soup = BeautifulSoup(r, 'html.parser')
    file = open('{}.txt'.format(i+1), 'w', encoding="utf-8")
    for article in soup.find_all(["h1", "div"], {"class": ["entry-title", "td-post-content"]}):
        art = article.text
        file.write(art)
    file.close()


# STOP WORD LIST
with open('D:\Python programs-2\StopWords_Generic.txt', 'r') as f:
    stop_words = f.read()

stop_words = stop_words.split('\n')

# MASTER DICTIONARY
master_dic = pd.read_excel('LoughranMcDonald_MasterDictionary_2020.xlsx')
positive_dictionary = [x for x in master_dic[master_dic['Positive'] != 0]['Word']]
negative_dictionary = [x for x in master_dic[master_dic['Negative'] != 0]['Word']]


def tokenize(text):
    text = re.sub(r'[^\w\s]', "", text)
    return word_tokenize(text)


def remove_stopwords(words, stop_words):
    return [x for x in words if not x.lower() in stop_words]


# Positive Score
def pos_score(store, words):
    numPosWords = 0
    for x in words:
        if x in store:
            numPosWords += 1
    return numPosWords


# Negative Score
def neg_score(store, words):
    numNegWords = 0
    for x in words:
        if (x in store):
            numNegWords -= 1
    sumNeg = numNegWords * -1
    return sumNeg


# Polarity Score
def polarity(positive_score, negative_score):
    pol_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    return pol_score


# Subjectivity Score
def subjectivity(positive_score, negative_score, num_words):
    return (positive_score + negative_score) / (num_words + 0.000001)


def syllable_morethan2(word):
    if len(word) > 2 and (word[-2:] == 'es' or word[-2:] == 'ed'):
        return False

    count = 0
    vowels = ['a', 'e', 'i', 'o', 'u']
    for i in word:
        if i.lower() in vowels:
            count = count + 1

    if (count > 2):
        return True
    else:
        return False


# Fog Index
def fog_index_cal(average_sentence_length, percentage_complexwords):
    return round((0.4 * (average_sentence_length + percentage_complexwords)),2)


# Sylable per word
def syllword(words):
    vowels = ['a', 'e', 'i', 'o', 'u']
    syllables=0
    for word in words:
        if word.endswith(('es', 'ed')):
            pass
        else:
            for w in word:
                if w.lower() in vowels:
                    syllables += 1
    return syllables


data=[]
for root, folders, files in os.walk("D:\Python programs-2\ASSIGNMENT"):
    for file in files:
        if file.endswith('.txt'):
            with open(os.path.join(root, file), 'r', encoding="utf-8") as f:
                text = f.read()
                data.append(text)


pos = []
neg = []
pol = []
sub = []
asl = []
pc = []
fi = []
aws = []
cwc = []
wc = []
sw = []
pp = []
awl = []
for content in data:
    tokenized_words = tokenize(content)
    words = remove_stopwords(tokenized_words, stop_words)
    num_words = len(words)
    wc.append(num_words)

    positive_score = pos_score(positive_dictionary, words)
    pos.append(positive_score)
    negative_score = neg_score(negative_dictionary, words)
    neg.append(negative_score)

    polarity_score = polarity(positive_score, negative_score)
    pol.append(polarity_score)

    subjectivity_score = subjectivity(positive_score, negative_score, num_words)
    sub.append(subjectivity_score)

    # Average Sentence Length
    sentences = sent_tokenize(content)
    num_sentences = len(sentences)
    average_sentence_length = round(num_words / num_sentences, 2)
    asl.append(average_sentence_length)

    # Complex Word Count
    num_complexword = 0
    for word in words:
        if syllable_morethan2(word):
            num_complexword = num_complexword + 1
    cwc.append(num_complexword)

    # %age of Complex Words
    percentage_complexwords = round((num_complexword / num_words)*100,2)
    pc.append(percentage_complexwords)

    # Fog Index
    fog_index = fog_index_cal(average_sentence_length, percentage_complexwords)
    fi.append(fog_index)

    # Average Number of Words per sentence
    Average_words_per_sentence = round((len(tokenized_words) / num_sentences),2)
    aws.append(Average_words_per_sentence)

    # Syllable per word
    syllable_per_word = syllword(words)
    sw.append(syllable_per_word)

    # Personal Pronouns
    personal_pronoun = 0
    pronouns = ['I', 'we', 'my', 'ours', 'us', 'We', 'My','Us']
    for word in words:
        if word in pronouns:
            personal_pronoun+=1

    pp.append(personal_pronoun)

    # Average Word Length
    num_chars = 0
    for word in words:
        num_chars += len(word)

    avg_word_length = round((num_chars / len(tokenized_words)),2)
    awl.append(avg_word_length)


url_list = range(1,171)
output_list = pd.DataFrame(
    {'URL_ID': url_list
     })

output_list['URL'] = urls
output_list['POSITIVE SCORE'] = pos
output_list['NEGATIVE SCORE'] = neg
output_list['POLARITY SCORE'] = pol
output_list['SUBJECTIVITY SCORE'] = sub
output_list['AVG SENTENCE LENGTH'] = asl
output_list['PERCENTAGE OF COMPLEX WORDS'] = pc
output_list['FOG INDEX'] = fi
output_list['AVG NUMBER OF WORDS PER SENTENCE'] = aws
output_list['COMPLEX WORD COUNT'] = cwc
output_list['WORD COUNT'] = wc
output_list['SYLLABLE PER WORD'] = sw
output_list['PERSONAL PRONOUNS'] = pp
output_list['AVERAGE WORD LENGTH'] = awl
output_list.set_index('URL_ID', inplace=True)
pd.set_option('display.max_columns', None)
print(output_list)

output_list.to_excel('Output.xlsx')