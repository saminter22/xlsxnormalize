import pandas as pd
import re


regex_email = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
regexp_phone = r'^((8|\+374|\+994|\+995|\+375|\+7|\+380|\+38|\+996|\+998|\+993)[\- ]?)?\(?\d{3,5}\)?[\- ]?\d{1}[\- ]?\d{1}[\- ]?\d{1}[\- ]?\d{1}[\- ]?\d{1}(([\- ]?\d{1})?[\- ]?\d{1})?$'


def normalize_phone(phone):
    try:
        phone = re.sub('\D', '', phone)
        phone = re.sub('^8', '7', phone)
        if len(phone) == 10:
            phone = '7' + phone
        if len(phone) != 11:
            return
        return phone
    except:
        return

def normalize_email(email):
    try:
        email = re.sub('\s', '', email)
        if not (re.fullmatch(regex_email, email)):
            email = ''
        return email.lower()
    except:
        return

def extract_domain(email):
    try:
        return email.split('@')[1]
    except:
        return


source_df = pd.read_excel('source/example.xlsx')
source_df['Дата регистрации'] = pd.to_datetime(source_df['Дата регистрации'])
source_df = source_df.sort_values(by=['Дата регистрации', 'Контактное лицо'], ascending=True)
proxy_df = source_df.copy(deep=False)
proxy_df = proxy_df.reset_index() 
proxy_df.to_excel (r'out/proxy_result.xlsx')
df = proxy_df[['Дата регистрации','Контактное лицо', 'Мобильный телефон', 'Телефон', 'Электронная почта']]
df['Телефон'] = df['Телефон'].apply(normalize_phone)
df['Мобильный телефон'] = df['Мобильный телефон'].apply(normalize_phone)
df = df[df['Мобильный телефон'].str[1].eq('9')]
df['Электронная почта'] = df['Электронная почта'].apply(normalize_email)
df['Домен сайта'] = df['Электронная почта'].apply(extract_domain)
df['Контактное лицо'] = df['Контактное лицо'].str.title()
result_df = df[['Дата регистрации','Контактное лицо', 'Мобильный телефон', 'Телефон', 'Электронная почта', 'Домен сайта']]
result_df.drop_duplicates('Электронная почта', keep='last', inplace=True)
result_df = result_df.reset_index()
result_df = result_df.drop('index', axis=1)

result_df.to_excel (r'out/result2.xlsx')
