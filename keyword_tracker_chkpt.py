# %%
import pandas as pd
import numpy as np
# from win32com import client
import os
import re
from nltk.corpus import stopwords
import nltk
import pickle

# !python -m pip install konlpy
# !python -m pip install JPype1

# %%
from konlpy.tag import Okt
from collections import Counter
import win32com.client

# %%
import importlib
import text_processing_utils
importlib.reload(text_processing_utils)
from text_processing_utils import TextProcessing

# %%
file_path = "wordlist_CP_EED.xlsx"
excel = win32com.client.Dispatch("Excel.Application")
excel.ScreenUpdating = True
file_dir = os.getcwd()
print(file_dir)
wb = excel.Workbooks.Open(file_dir + '\\' + file_path)
ws = wb.Sheets(1)

# sentences = [ws.Cells(row+1, 1).Value for row in range(1, ws.UsedRange.Rows.Count + 1)]
# labels = [ws.Cells(row+1, 6).Value for row in range(1, ws.UsedRange.Rows.Count + 1)]
# labels = labels[:-1]

data = []
for row in range(1, ws.UsedRange.Rows.Count + 1):
    cp = ws.Cells(row+1, 1).Value
    # print(cp)
    sub_mission = ws.Cells(row+1, 2).Value
    mission = ws.Cells(row+1, 3).Value
    # print(sub_mission)
    labels = ws.Cells(row+1, 5).Value
    causes = ws.Cells(row+1, 7).Value
    responsibility = ws.Cells(row+1, 9).Value
    delay_days = ws.Cells(row+1, 10).Value
    # labels = labels[:-1]
    data.append([cp, sub_mission, mission, labels, causes, responsibility, delay_days])

# wb.Close(False)
# excel.Quit()
# print(labels)
df = pd.DataFrame(data, columns=['cp', 'sub_mission', 'mission', 'labels', 'causes', 'responsibility', 'delay_days'])


# %%
df.head()
# print(data)

# %%
df.to_pickle('df_PM_EED.pkl')

# %%
df = pd.read_pickle('df_PM_EED.pkl')
df.head()

# %%


# %%
import importlib
import text_processing_utils
importlib.reload(text_processing_utils)
from text_processing_utils import TextProcessing

# Initialize tokenizer
okt = Okt()

# nltk.download('stopwords')
# stop_words = stopwords.words('english')
stop_words = TextProcessing.get_stop_words()
protected_phrases = TextProcessing.get_protected_phrases()

# Sort the phrases in descending order by length
protected_phrases = sorted(protected_phrases, key=len, reverse=True)

lot_regex_pattern = r'\b[A-Za-z0-9]{5}\.[1]\b'
replacement_word = 'LOTID'

# Build a regular expression pattern that matches any of the protected phrases
pattern = "|".join(re.escape(phrase) for phrase in protected_phrases)

all_tokens = []
master_tokens = []
# Process each sentence in the list

for index in df.index:
    # Convert to string in case any of the columns are not of string type
    cp = str(df.loc[index, 'cp']) if df.loc[index, 'cp'] is not None else ''
    sub_mission = str(df.loc[index, 'sub_mission']) if df.loc[index, 'sub_mission'] is not None else ''
    mission = str(df.loc[index, 'mission']) if df.loc[index, 'mission'] is not None else ''

    sentence = f"{mission} {sub_mission} {cp}"

# for sentence in sentences:
#     sentence = str(sentence)
    # Use re.finditer() to find all matches and replace them with placeholders
    for match in re.finditer(pattern, sentence):
        phrase = match.group()
        idx = protected_phrases.index(phrase)
        placeholder = f"PLACEHOLDER{idx:04}"
        sentence = sentence.replace(phrase, placeholder, 1)  # Replace only once to handle repeated phrases

    # Tokenize the transformed sentence
    tokenized_sentence = okt.morphs(sentence)

    # Merge PLACEHOLDER with its index in the token list
    merged_tokens = []
    idx = 0
    while idx < len(tokenized_sentence):
        token = tokenized_sentence[idx]
        if token == "PLACEHOLDER" and idx + 1 < len(tokenized_sentence) and tokenized_sentence[idx + 1].isdigit():
            merged_tokens.append(token + tokenized_sentence[idx + 1])
            idx += 2
        else:
            merged_tokens.append(token)
            idx += 1

    # Replace placeholders with their original phrases in the merged token list
    for idx, token in enumerate(merged_tokens):
        match = re.search(r"PLACEHOLDER(\d{1,4})", token)
        if match:
            placeholder_idx = int(match.group(1))
            merged_tokens[idx] = protected_phrases[placeholder_idx]

    # Remove stop words and append the tokens to all_tokens list
    filtered_tokens = [token for token in merged_tokens if token not in stop_words]
    all_tokens.append(filtered_tokens)

    # Extend the master_tokens list with the current sentence's tokens
    master_tokens.extend(filtered_tokens)

    df.at[index, 'token'] = ' '.join(filtered_tokens)

# print(all_tokens[:10])
# print(len(master_tokens))
# print(master_tokens)
df.tail()

# %%
word_freq = Counter(master_tokens)
top_keywords = word_freq.most_common(500)

# %%
print(top_keywords)

# %%
token_column = 11
# for i, (sentence, token_list) in enumerate(zip(sentences, all_tokens), start=1):
#     ws.Cells(i+1, token_column).Value = ', '.join(token_list)
    
# wb.SaveAs('wordlist_CP_EED_token_included.xlsx')
# # wb.Close()


for i in df.index:
    ws.Cells(i + 2, token_column).Value = df.loc[i, 'token']
wb.SaveAs('wordlist_CP_EED_token_included.xlsx')
# wb.Close()


# %%
import numpy as np
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import SVR
from sklearn.model_selection import GridSearchCV 
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.naive_bayes import MultinomialNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import classification_report, accuracy_score
from sklearn.metrics import mean_squared_error
from sklearn.metrics import confusion_matrix
from sklearn.metrics import classification_report
import pandas as pd

import matplotlib.pyplot as plt
import seaborn as sns


# %%
#Vectorize for test
vectorizer = TfidfVectorizer(analyzer="word", min_df=2, ngram_range=(1, 3), max_features=2000)

raw_sentences = [" ".join(tokens) for tokens in all_tokens]

X = vectorizer.fit_transform(df['token'][:-3])
y = np.array(df['delay_days'][:-3])
y_impact_ratio = np.array(df['labels'][:-3])


# %%
#Split data into training and test sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
X_train, X_test, y_impact_ratio_train, y_impact_ratio_test = train_test_split(X, y_impact_ratio, test_size=0.2, random_state=42)


# %%
#Train a linear SVR model
model = SVR(kernel='rbf', C=1000, gamma=1)
model.fit(X_train, y_impact_ratio_train)

#make predictions
predictions = model.predict(X_test)

# print("Mean Squared Error:", mean_squared_error(y_impact_ratio_test, predictions))
# mean_squared_error(y_true, Y-pred)

score = cross_val_score(model, X_test, predictions, cv=5)
print(score)
print(predictions)

# from pprint import pprint
# pprint(model.get_params())


# %%
param_grid = {'C': [0.1, 1, 10, 100, 1000],  
              'gamma': [1, 0.1, 0.01, 0.001, 0.0001], 
              'kernel': ['rbf']}  
  
grid = GridSearchCV(SVR(), param_grid, refit = True, verbose = 3) 
  
# fitting the model for grid search 
grid.fit(X_train, y_train) 
predictions = grid.predict(X_test)

results = grid.cv_results_
grid.best_params_
# results

# %%
delay_day_predict = model.predict(X)
delay_day_predict = np.clip(delay_day_predict, 0, None).round(0)

# %%
impact_ratio_predict = model.predict(X)
impact_ratio_predict = np.clip(impact_ratio_predict, 0, None).round(1) # clip negative values to 0

# %%
#Train Naive Bayes
nb_classifier = MultinomialNB()
nb_classifier.fit(X_train, y_train.astype('str'))
nb_predictions = nb_classifier.predict(X_test)

score = cross_val_score(nb_classifier, X_test, nb_predictions, cv=5)
print(classification_report(y_test.astype('str'), nb_predictions))
print(score)
print(nb_predictions)

impact_ratio_label_discrete = nb_classifier.predict(X)


# %%
#Train Random Forest model
rf_model = RandomForestRegressor()
rf_model.fit(X_train, y_train)
rf_predictions = rf_model.predict(X_test)

score = cross_val_score(rf_model, X_test, rf_predictions, cv=5)
print(score)
print(rf_predictions)

# %%
delay_day_predict_rf = rf_model.predict(X)
delay_day_predict_rf = np.clip(delay_day_predict_rf, 0, None).round(0)

# %%
df_token = pd.DataFrame(all_tokens[:-3])
# df_vector = pd.DataFrame(X.toarray())
df_token
# df_vector

# %%
df_token["impact_ratio_predict"] = impact_ratio_predict
# df_token["impact_ratio_predict_discrete"] = impact_ratio_label_discrete
df_token["delay_day_predict"] = delay_day_predict_rf
df_token["impact ratio"] = y_impact_ratio
df_token["delay_days"] = y
df_token["cp"] = df["cp"]
df_token["sub_mission"] = df["sub_mission"]
df_token["mission"] = df["mission"]
df_token = df_token.fillna('')
# df_analyzed = df_token.iloc[:,-3:]
df_analyzed = df_token[["mission", "sub_mission", "cp", "delay_days", "delay_day_predict", "impact ratio", "impact_ratio_predict"]]
df_analyzed
# df_sorted = df_analyzed.sort_values(by=['impact_ratio_predict'], ascending=False)
df_sorted = df_analyzed.sort_values(by=['impact_ratio_predict', 'delay_day_predict'], ascending=False)
df_sorted.to_csv("CPdata_delay_day_predicted.csv", index=False, encoding='utf-8-sig')
df_sorted

# %%

wb_final = excel.Workbooks.Add()
ws_final = wb_final.Worksheets(1)

excel.ScreenUpdating = True
# excel.Calculation = win32com.client.constants.xlCalculationManual

# try:
#     for col, column_name in enumerate(df_analyzed.columns, start=1):
#         ws.Cells(1, col).Value = column_name
#         for row, value in enumerate(df_analyzed.iloc[:, col-1], start=2):
#             try:
#                 ws.Cells(row, col).Value = value
#             except Exception as e:
#                 print(f"Error writing to cell ({row}, {col}): {value}")
#                 print(f"Error message: {e}")

# except Exception as e:
#     print("An unexpected error occurred:", e)

# finally:
#     # Restore Excel settings
#     excel.Calculation = win32com.client.constants.xlCalculationAutomatic
#     excel.ScreenUpdating = True

StartRow = 1
StartCol = 1

for col_num, col_name in enumerate(df_analyzed.columns, start=1):
    ws_final.Cells(StartRow, col_num).Value = col_name
    for row_num, value in enumerate(df_analyzed[col_name], start=2):
        ws_final.Cells(row_num, col_num).Value = value

wb_final.SaveAs('wordlist_CP_impact_ratio_prediction.xlsx')
# wb.Close(False)
# excel.Quit()
# print(labels)

# %%
impact_ratio_col = 7
impact_ratio_discrete_col = 8

for i in df_analyzed.index:
    ws.Cells(i + 2, impact_ratio_col).Value = df_analyzed.loc[i, 'impact_ratio_predict']
    ws.Cells(i + 2, impact_ratio_discrete_col).Value = df_analyzed.loc[i, 'impact_ratio_predict_discrete']


# for i, (sentence, token_list) in enumerate(zip(sentences, all_tokens), start=1):
#     ws.Cells(i+1, token_column).Value = ', '.join(token_list)
    
wb.SaveAs('wordlist_CP_token_impact_ratio_predict.xlsx')
# wb.Close()


# %%
#시각화
# mat = confusion_matrix(y_test, predictions)
mat = confusion_matrix(y_test.astype(str), nb_predictions.astype(str))
sns.heatmap(mat.T, square=True, annot=True, fmt='d', cbar=False) #xticklabels=train.target_names, yticklabels=train.target_names)

#Confusion Matrix Heatmap
plt.xlabel('true label')
plt.ylabel('predicted label')

# %%
#VOC 가져오기
df_VOC = pd.read_pickle('df_VOC.pkl')
df_VOC.head()

# %%
import importlib
import text_processing_utils
importlib.reload(text_processing_utils)
from text_processing_utils import TextProcessing


# Initialize tokenizer
okt = Okt()
# sentences

# nltk.download('stopwords')
# stop_words = stopwords.words('english')
stop_words = TextProcessing.get_stop_words()
protected_phrases = TextProcessing.get_protected_phrases()

# Sort the phrases in descending order by length
protected_phrases = sorted(protected_phrases, key=len, reverse=True)

# Build a regular expression pattern that matches any of the protected phrases
pattern = "|".join(re.escape(phrase) for phrase in protected_phrases)

voc_tokens = []
voc_master_tokens = []
# Process each sentence in the list

for index in df_VOC.index:
    # Convert to string in case any of the columns are not of string type
    project = str(df_VOC.loc[index, 'project']) if df_VOC.loc[index, 'project'] is not None else ''
    subject = str(df_VOC.loc[index, 'subject']) if df_VOC.loc[index, 'subject'] is not None else ''

    sentence = f"{project} {subject}"

# for sentence in sentences:
#     sentence = str(sentence)
    # Use re.finditer() to find all matches and replace them with placeholders
    for match in re.finditer(pattern, sentence):
        phrase = match.group()
        idx = protected_phrases.index(phrase)
        placeholder = f"PLACEHOLDER{idx:04}"
        sentence = sentence.replace(phrase, placeholder, 1)  # Replace only once to handle repeated phrases

    # Tokenize the transformed sentence
    tokenized_sentence = okt.morphs(sentence)

    # Merge PLACEHOLDER with its index in the token list
    voc_merged_tokens = []
    idx = 0
    while idx < len(tokenized_sentence):
        token = tokenized_sentence[idx]
        if token == "PLACEHOLDER" and idx + 1 < len(tokenized_sentence) and tokenized_sentence[idx + 1].isdigit():
            voc_merged_tokens.append(token + tokenized_sentence[idx + 1])
            idx += 2
        else:
            voc_merged_tokens.append(token)
            idx += 1

    # Replace placeholders with their original phrases in the merged token list
    for idx, token in enumerate(voc_merged_tokens):
        match = re.search(r"PLACEHOLDER(\d{1,4})", token)
        if match:
            placeholder_idx = int(match.group(1))
            voc_merged_tokens[idx] = protected_phrases[placeholder_idx]

    # Remove stop words and append the tokens to voc_tokens list
    voc_filtered_tokens = [token for token in voc_merged_tokens if token not in stop_words]
    voc_tokens.append(voc_filtered_tokens)

    # Extend the voc_master_tokens list with the current sentence's tokens
    voc_master_tokens.extend(voc_filtered_tokens)


# Assuming voc_tokens is a list of lists of tokens
combined_phrases = [' '.join(tokens) for tokens in voc_tokens]

# Add this as a new column to your DataFrame
# df_VOC['combined_phrases'] = combined_phrases
df_VOC.at[index, 'token'] = ' '.join(voc_filtered_tokens)

print(voc_tokens[:10])
# print(len(voc_master_tokens))
# print(voc_master_tokens)

# %%
#Vectorize for test
vectorizer = TfidfVectorizer(analyzer="word", min_df=2, ngram_range=(1, 3), max_features=2000)

raw_sentences = [" ".join(tokens) for tokens in voc_tokens]

x_voc = vectorizer.fit_transform(raw_sentences[:-1])

# %%
delay_day_predict_voc = rf_model.predict(x_voc)
delay_day_predict_voc = np.clip(delay_day_predict_voc, 0, None).round(0)

impact_ratio_predict_voc = model.predict(x_voc)
impact_ratio_predict_voc = np.clip(impact_ratio_predict_voc, 0, None).round(1) # clip negative values to 0

# %%
df_VOC_token = pd.DataFrame(voc_tokens[:-1])
# df_vector = pd.DataFrame(X.toarray())
df_VOC_token.tail()

# %%
df_VOC_token["impact_ratio_predict"] = impact_ratio_predict_voc
# df_VOC_token["impact_ratio_predict_discrete"] = impact_ratio_label_discrete
df_VOC_token["delay_day_predict"] = delay_day_predict_voc
df_VOC_token["subject"] = df_VOC["subject"]
df_VOC_token["project"] = df_VOC["project"]
df_VOC_token = df_VOC_token.fillna('')
# df_VOC_analyzed = df_VOC_token.iloc[:,-3:]
df_VOC_analyzed = df_VOC_token[["project", "subject", "delay_day_predict", "impact_ratio_predict"]]
df_VOC_analyzed
# df_VOC_sorted = df_VOC_analyzed.sort_values(by=['impact_ratio_predict'], ascending=False)
df_VOC_sorted = df_VOC_analyzed.sort_values(by=['impact_ratio_predict', 'delay_day_predict'], ascending=False)
df_VOC_sorted.to_csv("VOC_delay_day_impact_ratio_predicted.csv", index=False, encoding='utf-8-sig')
df_VOC_sorted


