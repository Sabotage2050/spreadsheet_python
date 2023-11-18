import streamlit as st
import requests
import json

page = st.sidebar.selectbox('Choose your page', ['registration', 'list'])

if page == 'registration':
    st.title('メモ登録画面')
    with st.form(key='registration'):
        content: str = st.text_input('メモ内容', max_chars=100)
        data = {
                'content': content
        }
        submit_button = st.form_submit_button(label='メモ登録')

        if submit_button:
            url = 'http://127.0.0.1:8000/memos'
            res = requests.post(
                url,
                data=json.dumps(data)
            )
            if res.status_code == 200:
                st.success('メモ登録完了')
            st.json(res.json())

elif page == 'list':
    st.title('メモ一覧画面')
    res = requests.get('http://127.0.0.1:8000/memos')
    records = res.json()
    print(records)
    for record in records:
        st.subheader('・' + record.get('content'))