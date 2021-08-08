# -*- coding: utf-8 -*-
"""
Created on Thu Jul 29 06:41:28 2021

@author: dharm
"""

import streamlit as st
import pandas as pd
import os
import base64
from list_content import LessonLearn


def search_relevant_file (nshow, text):
    from nltk.tokenize import word_tokenize
    from nltk.corpus import stopwords
    from string import punctuation
    
    #menggunakan stopword indonesia untuk remove
    sw_indo = stopwords.words("indonesian") + list(punctuation)
    
    #membuka dokumen dan set variable yang ingin di search
    df = pd.read_csv("D:\\0. python\\01. streamlit\\data\\index_lesson_learned_data.csv")
    target = df.clean_summary
    
    #import module sklearn untuk Tfidf
    from sklearn.feature_extraction.text import TfidfVectorizer, TfidfTransformer
    
    #fit dan transform
    tfidf = TfidfVectorizer(ngram_range = (1,2), tokenizer = word_tokenize, stop_words=sw_indo)
    tfidf_matrix = tfidf.fit_transform (target)
    
    #import cosine similarity
    from sklearn.metrics.pairwise import cosine_similarity
    
    # argsort untuk show index dari
    sim = cosine_similarity(tfidf_matrix[0], tfidf_matrix)
    sim.argsort()
    
    # argsort untuk show index dari
    # mengubah ke lower
    text = text
    text = text.lower()
    
    #mengubah text menjadi tfidf, perlu di buat dulu matrix nya
    #kemudian di transform
    tfidf_query = tfidf.fit(target)
    tfidf_transform = tfidf_query.transform([text])
    
    #tfidf_matrix_to_search = tfidf.transform(teks)
    sim = cosine_similarity(tfidf_transform, tfidf_matrix)
    to_show = sim.argsort()
    
    #reversed order matrix dari yang paling relevant menadi tidak relevant
    show = list (reversed (to_show [0][-nshow:]))
    return df.loc[show, ['directory', 'clean_summary']].to_dict(orient='index')

def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream; base64,{bin_str}" download="{os.path.basename(bin_file)}" style="font-size: 10px">Download {file_label}</a>'
    return href

def upload_file(path):
    '''
    A function to upload a file and return the its filename
    '''
    file_name = ""
    uploaded_files = st.file_uploader("Choose a file", accept_multiple_files=True)
    upload = st.button("Upload")
    if uploaded_files is not None and upload:
        try:
            for uploaded_file in uploaded_files:
                file_loct = path + uploaded_file.name
                with open(file_loct,'wb') as out:
                    out.write(uploaded_file.read())
                st.success(uploaded_file.name + " is uploaded!") #indicate the file has been uploaded
                file_name = uploaded_file.name
                uploaded_file.close()
        except:
            pass
    return file_name


page_options = ["HOME", "UPLOAD DATA"]
page = st.sidebar.selectbox("Section", options = page_options)


if page == "UPLOAD DATA":
    save_path = "D:\\0. python\\01. streamlit\\data\\lesson_learned\\"
    upload_file(save_path)
    LL= LessonLearn()
    list_files = LL.get_contents(save_path)
    list_files.to_csv("D:\\0. python\\01. streamlit\\data\\index_lesson_learned_data.csv")

if page == "HOME":
    st.markdown(f'<p style="font-size: 24px; color: #4285F4; margin-bottom: 0em;  margin-top: 0em;">GOOLL</p>', unsafe_allow_html=True)
    qry = st.text_input("On the go lesson learn")
    s1,s2,s3,s4,s5,s6 = st.beta_columns(6)
    
    n = s1.selectbox("show top result", [5,10,20])
    if qry != "":
        qry_result = search_relevant_file(n, qry)
        for i in qry_result:
            st.markdown(f'<p style="font-size: 12px; color: #4285F4; margin-bottom: 0em;  margin-top: 0em;">{qry_result[i]["directory"]}</p>'
                        f'<p style="font-size:12px; margin-bottom: 0em;  margin-top: 0em;">{qry_result[i]["clean_summary"][200:600]}</p>'
                        , unsafe_allow_html=True)
            if st.button(">>>", key = i):
                st.markdown(get_binary_file_downloader_html("D:\\0. python\\01. streamlit\\data\\lesson_learned\\"+qry_result[i]["directory"], qry_result[i]["directory"]), unsafe_allow_html=True)
            