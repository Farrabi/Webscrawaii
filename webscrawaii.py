import pandas as pd
import re
from time import sleep
import random
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk

def get1():
    global LETTER
    LETTER=[]
    for entry in Entries:
        LETTER.append(entry.get())
    

def L(data_):
    global N_URLS
    N_URLS=[]
    for c in range(len(data_)):
        if data_.iloc[c][:4]!='http':
            pass
        else:
            N_URLS.append(data_.iloc[c])
            DIC={ i : None for i in N_URLS}
    return  DIC

def scrape(url):
    try:
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('start-maximized')
        chrome_options.add_argument('disable-infobars')
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36")
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(url)
        sleep(3)
        innerHTML = driver.execute_script("return document.body.innerHTML")
        page_text=str(innerHTML)
        auto=re.findall(REGEX,page_text)
##        new_code.append(auto)
##        conn_urls.append(url)
        driver.quit()
    except Exception as e:
        Manu[str(url)]='Unables to connect need manual check'
        print(e)
        auto='An error occured please check url manually'
    except OSError as ee:
        Manu[str(url)]='Unables to connect need manual check'
        print(ee)
        auto='An error occured please check url manually'
    except IOError as eee:
        Manu[str(url)]='Unables to connect need manual check'
        print(eee)
        auto='An error occured please check url manually'
    else:
        print("connection successful")
    return auto


def compare(D1, D2):
    global N1
    N1=[]
    global N2
    N2=[]
    global N3
    N3=[]
    for D1_keys, D2_keys in zip(D1.keys(), D2.keys()):
        if D1_keys == D2_keys and D1[D1_keys]==D2[D2_keys]:
            N1.append(D2_keys)
            N2.append(D2[D2_keys])
            N3.append('OK')
            print('Ok', D1[D1_keys], D2[D2_keys])
        else:
            N1.append(D2_keys)
            N2.append(D2[D2_keys])
            N3.append('NOT OK')
            print('Not', D1[D1_keys], D2[D2_keys])



if __name__ == "__main__":
    code_letter=['自動車 自動車 及び バイク保険 バイク保険', 'ALセット版普通傷害保険', 'ALセット版交通傷害保険', 'ダイレクト傷害保険(普傷 ・家傷) ', 'グループ傷害保険(普傷 ・交普傷 ・交)', '入院手術保険', 'ペット保険']
    Entries=[]
    root_2=tkinter.Tk()
    root_2.geometry("720x406")
    root_2.title('Webscrawaii')
    image = Image.open("broly.jpg")
    background_image=ImageTk.PhotoImage(image)
    background_label = tkinter.Label(root_2, image=background_image)
    background_label.pack()
    y_=70
    i=70
    CODE=[]
    REGEX=''


    for letter in code_letter:
        label =  tkinter.Label(root_2, text = letter).place(x = 20,y = y_)
        y_=y_+20
        entry=tkinter.Entry(root_2)
        entry.place(x=600,y=i)
        i=i+20
        Entries.append(entry)

    label1 =tkinter.Label(root_2, text = 'Input first letter of the code. お願いね')
    label1.place(x =180,y =20)
    label1.config(font=('Calibri', 22))
    sbmitbtn = tkinter.Button(root_2, text = "Submit",activebackground = "pink", activeforeground = "blue", command=lambda: [get1(), root_2.destroy()])
    sbmitbtn.pack()

    sbmitbtn.place(x = 30, y = 220)


    root_2.mainloop()



    for c in LETTER:
        if len(c)!=0:
            CODE.append(c)

    REGEX=''
    for j in CODE:
        REGEX= j+'[0-9]{4,11}\-.{8}?'+'|'+j+'[0-9]{4,11}.?'+'|'+REGEX

    REGEX=REGEX[:-1]




    root = tkinter.Tk()
    root.withdraw()

    file = filedialog.askopenfile(parent=root,mode='rb',title='Choose a file')
    if file != None:
        data = file.read()
        file.close()
        print("I got %d bytes from this file." % len(data))



    df=pd.read_excel(file.name, sheet_name='URL一覧', encoding = 'utf-8')
    URLS=df.loc[:,'新URL']
    URLS_clean = URLS.fillna('no url')
    old_code= df.loc[:,'募文番号']
    code_clean=old_code.fillna('no code')
    code_clean=code_clean[2:]
    conn_urls=[]
    new_code=[]
    prox_list=['85.196.183.162:80', '202.182.121.205:80', '45.77.135.170:80', '54.180.123.253:8080']
    proxies={'https':random.choice(prox_list)}
    sessions=[]
    Manu={}

    new_origin={}

    Origin=pd.Series(code_clean.values, index=list(L(URLS_clean).keys())).to_dict()


    new_origin={k:scrape(k)for k , _ in L(URLS_clean).items()}



    #new_code=map(scrape, L(URLS_clean))



    ##for i in L(URLS_clean):
    ##    scrape(i)
    ##    print(i)


    df2=pd.DataFrame(data=list(new_origin.values()))
    df3=pd.DataFrame.from_dict(Manu, orient='index')


    for i in range(int(df2.shape[0])):
        for j in range(int(df2.shape[1])):
            if df2[j][i]==None:
                pass
            elif len(df2[j][i])<4:
                del df2.iloc[i,j]
            elif df2[j][i][-1].isdigit()==False:
                df2[j][i]=df2[j][i][:-1]
            else:
                pass

    codes=df2.values.tolist()

    for i, j in zip(new_origin,codes):
        new_origin[i]=j



    compare(Origin, new_origin)


    D= pd.DataFrame(list(zip(N1, N2, N3)), columns=['URLS', 'codes', 'check'])


    D.to_excel('new_codes.xlsx', encoding='utf_8_sig')
    df3.to_excel('URLS_to_be_checked_manually.xlsx', encoding='utf_8_sig')




