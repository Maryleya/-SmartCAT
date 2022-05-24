import time
import pandas as pd
from itertools import groupby
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager

useragent = UserAgent()
profile = webdriver.FirefoxProfile()
profile.set_preference("general.useragent.override", useragent.random)

def terms(needed_term):
    driver = webdriver.Firefox(executable_path = '')
    url = 'https://www.multitran.com/m.exe?l1=1&l2=2'
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    needed_term = needed_term.lower()
    
    links_terms = soup.find_all('a')
    full_links_terms = list(links_terms[10:])[:-2]
    entries = soup.find_all('td', {'align': 'right'})
    
    entriess = []
    for i in range(2, len(entries)-1):
        entriess.append(entries[i].contents[0])
    
    int_entries = []
    for i in entriess:
        int_entries.append(int(i.replace('.', '')))
    
    dict_of_links = {}
    entries_dict = {}
    for i in range(len(full_links_terms)):
        link = full_links_terms[i].get('href')
        term = full_links_terms[i].contents[0]
        term = term.lower().replace(u'\xa0', u' ')
        dict_of_links.update({term: link})
        entries_dict.update({term: int_entries[i]})
        
    driver.close()
    
    if needed_term in dict_of_links.keys():
        needed_link = dict_of_links[needed_term]
        needed_url = 'https://www.multitran.com{0}'.format(needed_link)
        needed_entry = entries_dict[needed_term]
        parse_nth_page(needed_url, needed_entry)
    else:
        print('No such category')

def parse_nth_page(url, entry):
    driver = webdriver.Firefox(executable_path = '')
    driver.get(url)
    rus_list = []
    eng_list = []
    while entry > 0:
        soup = BeautifulSoup(driver.page_source, 'html.parser')
    
        links_needed_terms = soup.find_all('td', {'class': 'termsforsubject'})
        full_links_needed_terms = list(links_needed_terms)
    
        for j in range(1, len(full_links_needed_terms), 3):
            try:
                rus_list.append(full_links_needed_terms[j].find('a').contents[0])
            except:
                rus_list.append('-')

        for j in range(0, len(full_links_needed_terms), 3):
            try:
                eng_list.append(full_links_needed_terms[j].find('a').contents[0])
            except:
                eng_list.append('-')
            entry = entry - 1
        
        l = driver.find_element_by_link_text('>>')
        l.click()
        time.sleep(2)
    
    driver.close()
    
    eng_rus_dict = {}
    for i in range(len(eng_list)):
        if eng_list[i] not in eng_rus_dict:
            eng_rus_dict[eng_list[i]] = [rus_list[i]]
        else:
            eng_rus_dict[eng_list[i]].append(rus_list[i])

    final_eng_rus_dict = {}
    for k, v in eng_rus_dict.items():
        new_v = [el for el, _ in groupby(v)]
        final_eng_rus_dict[k] = new_v

    eng_list = list(final_eng_rus_dict.keys())
    rus_list = list(final_eng_rus_dict.values())
    
    new_rus_list = []
    for i in rus_list:
        new_i = set(i)
        new_i = str(new_i).strip('}][{')
        new_i = new_i.replace("'", '')
        new_rus_list.append(new_i)
    
    write_to_xlsx(new_rus_list, eng_list)

def write_to_xlsx(rus_list, eng_list):
    rus = {'ru Term1': rus_list}
    eng = {'en Term1': eng_list}
    
    d_keys = ['Comments', 'CreationDate', 'Author', 'LastModifiedDate', 'LastModifiedBy', 'en Term1', 'ru Term1']
    eng_to_rus = dict.fromkeys(d_keys)
    eng_to_rus.update(eng)
    eng_to_rus.update(rus)
    
    df = pd.DataFrame(eng_to_rus)
    df = df.loc[df['en Term1'] != '-']
    df = df.loc[df['ru Term1'] != '-']
    
    name = 'eng_to_rus.xlsx'
    df.to_excel(name, index=None)

term = ''
terms(term)