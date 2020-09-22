

import pandas as pd
import csv
import xlwt 
from xlwt import Workbook 
import gensim
import gensim.corpora as corpora
from gensim.models import CoherenceModel
from nltk.corpus import stopwords
import matplotlib.pyplot as plt
import nltk
from nltk.stem import PorterStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
from gensim.corpora import Dictionary

#starting

#Dataset importing 
nod = int(input("Enter no. of documents for minning:"))
notp = int(input("Enter no. of topics:"))
pr = float(input("Enter the perecentage of words to be removed in decimal points:"))


df = pd.read_excel (r'dataset.xlsx')
data=[]

for i in range(nod):
    data.append(df.iloc[i, 0]) 
#Function data as argument/parameter 
def pos(t):
    words = set(nltk.corpus.words.words())
    TEXT = t
    #task-2 removing non English words
    After_removal_non_english = TEXT 
    After_removal_non_english  = " ".join(w for w in nltk.wordpunct_tokenize(After_removal_non_english ) if w.lower() in words or not w.isalpha())
    
    #task-1 tokenizing words and removing punctutation
    tokenizer = nltk.RegexpTokenizer(r"\w+")
    words = tokenizer.tokenize(After_removal_non_english)
    #task-3 dividing into parts of speech
    pos=nltk.pos_tag(words)
    nouns=[]
    adjectives=[]
    adverbs=[]
    list_of_ana=[]
    for i in range(len(pos)):
        if (pos[i][1]== 'NNP' or pos[i][1]== 'NN' or pos[i][1]== 'NNS' or pos[i][1]== 'NNPS'):
            nouns.append(pos[i][0])
        if (pos[i][1]== 'JJ' or pos[i][1]== 'JJR' or pos[i][1]== 'JJS'):
            adjectives.append(pos[i][0])
        if (pos[i][1]== 'RB' or pos[i][1]== 'RBR' or pos[i][1]== 'RBS'):
            adverbs.append(pos[i][0])
     
        for i in range(len(pos)):
            if (pos[i][1]== 'NNP' or pos[i][1]== 'NN' or pos[i][1]== 'NNS' or pos[i][1]== 'NNPS'or pos[i][1]== 'JJ' or pos[i][1]== 'JJR' or pos[i][1]== 'JJS' or pos[i][1]== 'RB' or pos[i][1]== 'RBR' or pos[i][1]== 'RBS'):
                list_of_ana.append(pos[i])
        
   
    Combined_Pos=nouns+adjectives+adverbs
    
    return pos , Combined_Pos , nouns , adjectives , adverbs , list_of_ana



nouns=[]
adverbs=[]
adjectives=[]
full_review=[]
aan_review=[]
full_list_of_ana=[]
for i in range(nod):
    fpos, cpos , nn , aj , ad , fla =pos(data[i])
    nouns.append(nn)
    adverbs.append(aj)
    adjectives.append(ad)
    full_review.append(fpos)
    aan_review.append(cpos)
    full_list_of_ana.append(fla)

csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Tagged_corpus.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(full_review)
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Nouns.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(full_list_of_ana)
#printing in a file   fpos cpos

csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('All_Tags.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(full_review)
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Nouns.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(nouns)
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Adjectives.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(adjectives)
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Adverbs.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(adverbs)



def t4_5(c):
    #task-4 removing stop words
    stop_words = set(stopwords.words('english'))
    filtered_sentence = [w for w in c if not w in stop_words]
    filtered_sentence = []
    for w in c:
        if w not in stop_words:
            filtered_sentence.append(w)
    
    #task-5 stemming
    ps = PorterStemmer()
    stemmed=[]
    for w in filtered_sentence:
        stemmed.append(ps.stem(w))
    return stemmed


stemmed_all=[]
for i in range(nod):
    sa=t4_5(aan_review[i])
    stemmed_all.append(sa)



texts=stemmed_all
#from gensim.corpora import Dictionary
dictionary  = Dictionary(stemmed_all)

dictionary .filter_extremes(no_below=pr, keep_n=None)

  
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('dict') 
for i in range(len(dictionary)):
    sheet1.write(i+1, 0, dictionary[i]) 

wb.save('Dictionary_after_2%_words_removal.xls') 

# remov%
#all_tokens = sum(texts, [])
#tokens_once = set(word for word in set(all_tokens) if all_tokens.count(word) < (len(dictionary))*0.02)
#texts = [[word for word in text if word not in tokens_once]
#        for text in texts]
#stemmed_all=texts
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Stemmed_documents.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(stemmed_all)
mydict = corpora.Dictionary()
#dtm
doc_term_matrix = [dictionary.doc2bow(doc) for doc in stemmed_all]
mycorpus = [mydict.doc2bow(doc, allow_update=True) for doc in stemmed_all]
csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Document_term_matrix.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(doc_term_matrix)
#print(mycorpus)

word_counts = [[(mydict[id], count) for id, count in line] for line in mycorpus]

texts = stemmed_all

texts=stemmed_all
data_lemmatized=stemmed_all
id2word = corpora.Dictionary(data_lemmatized)

corpus_mapping=[]
corpus_presentation=[]

for i in range(nod):
    vectorizer = TfidfVectorizer(max_df=0.5)
    X = vectorizer.fit_transform(stemmed_all[i])
    corpus_mapping.append(X)
    corpus_presentation.append(vectorizer.get_feature_names())

corpus = [id2word.doc2bow(text) for text in texts]
lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                           id2word=id2word,
                                           num_topics=notp, 
                                           random_state=100,
                                           update_every=1,
                                           chunksize=10,
                                           passes=10,
                                           per_word_topics=True,
                                           alpha='asymmetric', 
                                           minimum_probability=1e-8)

#pprint(lda_model.print_topics())
doc_lda = lda_model[corpus]



print('\nPerplexity: ', lda_model.log_perplexity(corpus))  # a measure of how good the model is. lower the better.

coherence_model_lda = CoherenceModel(model=lda_model, texts=stemmed_all, dictionary=dictionary, coherence='u_mass')
coherence_lda = coherence_model_lda.get_coherence()
print('\nCoherence Score: ', coherence_lda)
fifty_terms=[]
for i,topic in lda_model.show_topics(formatted=True, num_topics=notp, num_words=50):
    fifty_terms.append(topic)
    #print()

csv.register_dialect('myDialect',
                     delimiter=' ')
with open('Fifty_Terms_of_each_topic.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(fifty_terms)

bow = []
for t in stemmed_all:
    bow.append(dictionary.doc2bow(t))
   
#tfidf = models.TfidfModel(bow)
#tf_matrix=[]
#for i in range(nod):
#    tf_matrix.append(tfidf[bow[i]])
#
#
##print(tf_matrix[7])
#
#
#def top_term(k):
#    tf_obj = tfidf[bow[k]]
#    z=sorted(tf_obj, key=lambda x: x[1], reverse=True)[:5]
#    return z
#
#t_terms_all=[]
#
#for i in range(nod):
#    l=[]
#    l=top_term(i)
#    t_terms_all.append(l)



#from gensim.models import ldamodel

    
all_topics= lda_model.get_document_topics(corpus, per_word_topics=True)
dtopics=[]
for doc_topics, word_topics, phi_values in all_topics:
    dtopics.append(doc_topics)

csv.register_dialect('myDialect',
                     delimiter=',',
                     quoting=csv.QUOTE_ALL)
with open('Topic Distribution.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='myDialect')
    writer.writerows(dtopics)




limit=80; start=1; step=12;
prep_all=[]
for num_topics in range(start, limit, step):
    lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                               id2word=id2word,
                                               num_topics=num_topics, 
                                               random_state=100,
                                               update_every=1,
                                               chunksize=10,
                                               passes=10,
                                               alpha='auto',
                                               per_word_topics=True)
    prep_all.append(lda_model.log_perplexity(corpus))  # a measure of how good the model is. lower the better.
x=range(start, limit, step)
plt.plot(x, prep_all)
plt.xlabel("Num Topics")
plt.ylabel("Preplexity")
plt.legend(("Preplexity_values"), loc='best')
plt.savefig('Preplexity_values.png', dpi=300, bbox_inches='tight')
plt.show()
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Preplexity_scores') 
for i in range(len(prep_all)):
    sheet1.write(i+1, 0, prep_all[i]) 

wb.save('Preplexity_scores.xls') 
limit=80; start=1; step=12;
coher_all=[]
for num_topics in range(start, limit, step):
    lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                               id2word=id2word,
                                               num_topics=num_topics, 
                                               random_state=100,
                                               update_every=1,
                                               chunksize=10,
                                               passes=10,
                                               alpha='auto',
                                               per_word_topics=True)
    
    prep_all.append(lda_model.log_perplexity(corpus))  # a measure of how good the model is. lower the better.
    coherence_model_lda = CoherenceModel(model=lda_model, texts=stemmed_all, dictionary=dictionary, coherence='u_mass')
    coherence_lda = coherence_model_lda.get_coherence()
    coher_all.append(coherence_lda)

x=range(start, limit, step)
plt.plot(x, coher_all)
plt.xlabel("Num Topics")
plt.xlabel("Num Topics")
plt.ylabel("Coherence score")
plt.legend(("coherence_values"), loc='best')
plt.savefig('coherence_values.png', dpi=300, bbox_inches='tight')
plt.show()
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Coherence score') 
for i in range(len(coher_all)):
    sheet1.write(i+1, 0, prep_all[i]) 

wb.save('Coherence score.xls') 