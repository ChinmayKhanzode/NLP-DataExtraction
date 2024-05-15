import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize,sent_tokenize
from nltk.corpus import cmudict
import openpyxl


variable_data = [["URL_ID","URL","POSITIVE SCORE","NEGATIVE SCORE","POLARITY SCORE","SUBJECTIVITY SCORE","AVG SENTENCE LENGTH",
                  "PERCENTAGE OF COMPLEX WORDS","FOG INDEX","AVG NUMBER OF WORDS PER SENTENCE","COMPLEX WORD COUNT","WORD COUNT","SYLLABLE PER WORD","PERSONAL PRONOUNS","AVG WORD LENGTH"]]


df = pd.read_excel('Input.xlsx')

if not os.path.exists('articles'):
    os.makedirs('articles')

def get_article():
    print("Loading.....")
    for index, row in df.iterrows():
        url = row['URL']
        url_id = row['URL_ID']


        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        title = soup.find('title').get_text()
        title = title.replace(" - Blackcoffer Insights", "")
        text = ' '.join([p.get_text() for p in soup.find_all('p')])


        p= soup.findAll(['p'])
        start = 0
        final_text = ""
        for ps in p:
            if(ps.get_text() == "We provide intelligence, accelerate innovation and implement technology with extraordinary breadth and depth global insights into the big data,data-driven dashboards, applications development, and information management for organizations through combining unique, specialist services and high-lvel human expertise."):
                break
            if(start == 1):
                final_text += ps.get_text() + "\n"
            if(ps.get_text() == "Code Review Checklist"):
                start = 1
        print(url_id)


        # Save the article in a text file
        with open(f'articles/{url_id}.txt', 'w', encoding='utf-8') as f:
            f.write(title + '\n' + final_text)

    print("Done!")

def get_all_stop_words():
    file = open('MasterDictionary/positive-words.txt','r')
    import os
    directory = "StopWords"
    combined_text = ""
    for filename in os.listdir(directory):
        if filename.endswith(".txt"):
            with open(os.path.join(directory, filename), 'r') as file:
                combined_text += file.read()
    return word_tokenize(combined_text)


def get_positive_score(content):

    file = open('MasterDictionary/positive-words.txt','r')
    pos_words = file.read()
    positive_dict = word_tokenize(pos_words)

    words = word_tokenize(content)
    positive_score = sum(word in positive_dict for word in words)

    print("Positive Score:", positive_score)
    file.close()
    return positive_score

def get_negative_score(content):
    file = open('MasterDictionary/negative-words.txt','r')
    neg_words = file.read()
    negative_dict = word_tokenize(neg_words)

    words = word_tokenize(content)
    negative_score = sum(word in negative_dict for word in words)

    print("Negative Score:", negative_score)
    file.close()
    return negative_score

def get_subjectivity_score(content,positive_score,negative_score,stop_words):
    for word in stop_words:
        content = content.replace(word, '')
    subjectivity_score = (positive_score+ negative_score)/ (len((content.split(" "))) + 0.000001)
    print("Subjectivity Score:", subjectivity_score)
    return subjectivity_score

def avg_sentence_len(content):
    num_of_sentences = len(content.split("."))
    num_of_words = len(content.split())
    avg_len = num_of_words/num_of_sentences
    print("Average Length Of Sentences:",avg_len)
    return avg_len


def get_fog_index(comp_word_count,content):
    length_of_words = len(word_tokenize(content))
    fog_index = 0.4*((length_of_words/len(content.split("."))) + (100 *(comp_word_count/length_of_words)))
    print("FOG Index:",fog_index)
    return fog_index

def avg_num_words_sent(content):
    sentences = nltk.sent_tokenize(content)
    word_counts = [len(nltk.word_tokenize(sentence)) for sentence in sentences]
    print("Average_Number of Words Per Sentence:",sum(word_counts) / len(sentences))
    return sum(word_counts) / len(sentences)



def word_count(content):
    word_count = len(word_tokenize(content))
    print("Word Count:",word_count)
    return word_count

def syllable_count_word(word):
    word = word.lower()
    vowels = "aeiouy"
    count = 0

    if word[0] in vowels:
        count += 1

    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1

    if word.endswith("e"):
        count -= 1

    if word.endswith("le"):
        count += 1

    if word.endswith("es") or word.endswith("ed"):
        count -= 1

    if count == 0:
        count += 1
    
    return count

def avg_syllable_word(content):
    words = word_tokenize(content)
    counter = []
    for word in words:
        count = syllable_count_word(word)
        counter.append(count)
    avg = sum(counter)/len(counter)
    print("Average Syllable Count:",avg)
    return avg
              

def personal_pronoun(content):
    pronouns = ["I", "we", "my", "ours", "us"]
    count= 0
    words = word_tokenize(content)
    for word in words:
        if word in pronouns and word != "US":
            count+=1
    print("Number of Personal Pronouns:", count)
    return count

def avg_word_len(content):
    words = word_tokenize(content)
    num_char = sum(len(word) for word in words)
    print("Average Word Length", num_char/len(words))
    return num_char/len(words)

def complex_word_count(content):
    # more than two syllables
    words = word_tokenize(content)
    count=0
    for word in words:
        if syllable_count_word(word)>2:count+=1
    print("Complex Word Count:",count)
    return count


def get_variables():
    print("getting variables")
    url_id_data = df['URL_ID'].tolist()
    url_data = df['URL'].tolist()
    stop_words = get_all_stop_words()
    
    for index,id in enumerate(url_id_data):
        file = open(f'articles/{id}.txt', 'r', encoding="utf8")
        print(index)
        content = file.read()
        
        positive_score = get_positive_score(content)
        negative_score = get_negative_score(content)
        polarity_score = (positive_score - negative_score)/ ((positive_score + negative_score) + 0.000001)
        print("Polarity Score:",polarity_score)
        subjectivity_score = get_subjectivity_score(content,positive_score,negative_score,stop_words)
        avg_sentence_length = avg_sentence_len(content)
        
        average_number_of_words = avg_num_words_sent(content)
        comp_word_count = complex_word_count(content)
        per_comp_word = (comp_word_count/len(word_tokenize(content)))*100
        print("Percentage Complex Word Count:", per_comp_word)
        fog_index = get_fog_index(comp_word_count,content)
        word_counter = word_count(content)
        syll_word = avg_syllable_word(content)
        per_pronouns = personal_pronoun(content) 
        avg_word_length = avg_word_len(content)

        file.close()
        variable_data.append([url_id_data[index],url_data[index],positive_score,negative_score,polarity_score,subjectivity_score,
                              avg_sentence_length,per_comp_word,fog_index,average_number_of_words,
                              comp_word_count,word_counter,syll_word,per_pronouns,avg_word_length])

def upload_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active

    for row in variable_data:
        ws.append(row)

    wb.save('output.xlsx')
    
get_article()
get_variables()
upload_to_excel()

