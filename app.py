import requests
import base64
import codecs
import time
import json
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import openpyxl
from datetime import datetime
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


key = 2
#wb = openpyxl.load_workbook('testfile.xlsx')
wb = openpyxl.Workbook()
worksheet = wb['Sheet'] #Делаем его активным
worksheet['A1']='Регион'
worksheet['B1']='Район'
worksheet['C1']='Город'
worksheet['D1']='Улица'
worksheet['E1']='Дом'
worksheet['F1']='Квартира'
worksheet['G1']='Этаж'
worksheet['H1']='Кадастровый номер'
worksheet['I1']='Дата присвоения кадастрового номера'
worksheet['J1']='Кадастровая стоимость (руб)'
worksheet['K1']='Дата определения'
worksheet['L1']='Дата внесения'
worksheet['M1']='Вид, номер государственной регистрации права'
worksheet['N1']='Дата государственной регистрации права'
worksheet['O1']='Ограничение прав и обременение объекта недвижимости'
worksheet['P1']='Дата ограничения прав'
#В указанную ячейку на активном листе пишем все, что в кавычках
wb.save('test.xlsx') #Сохраняем измененный файл

def get_captcha(ses, capKey):
    res = ses.get('https://lk.rosreestr.ru/account-back/captcha.png', verify=False)
    image = base64.encodebytes(res.content)
    url = 'https://rucaptcha.com/in.php'
    params = dict(key=capKey, method='base64', body=image, json=1)
    res = requests.post(url, params)
    #print(res.content)
    captcha = ''
    url = 'https://rucaptcha.com/res.php'
    params = dict(key=capKey, action='get', id=res.json()['request'], json=1)
    while True:
        time.sleep(1)
        res = requests.get(url, params)
        if int(res.json()['status']) == 1:
            # тут делать, что нужно, т.е. повторно отправлять запрос с решенной капчей
            # решенная капча в res.json()['request']
            #print(res.content)
            captcha = res.json()['request']
            break
        elif res.json()['status'] != '1':
            continue
        else:
            #print('ERROR')
            #print(res.json())
            break
    url = 'https://lk.rosreestr.ru/account-back/captcha/' + captcha
    res = ses.get(url, verify=False)
    #print(res)
    return captcha

captchakey = ''
with open('key.txt', encoding='utf-8') as f:
    captchakey = f.readline()

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36 OPR/83.0.4254.46'}
term = ''
with open('input.txt', encoding='utf-8') as f:    
    while True:
        term = f.readline()
        s = requests.Session()
        quer = term.strip()
        #quer = quer.replace('\n', '')
        last = quer.rfind(',')
        print(quer[:last])
        #print(quer)
        if not term:
            break
        url = f'https://lk.rosreestr.ru/account-back/address/search?term={quer}'
        res = s.get(url, verify=False)
        if quer.rfind('-') > 0:
            print("Нашел")
            a = int(quer[quer.rfind('№') + 1:quer.rfind('-')])
            b = int(quer[quer.rfind('-') + 1:])
            print(a)
            print(b)
            for j in range(a, b + 1):
                query = quer[:quer.rfind('№') + 1] + str(j)
                print(query)
                url = f'https://lk.rosreestr.ru/account-back/address/search?term={query}'
                res = s.get(url, verify=False)
                for i in res.json():
                    astring = i['full_name']
                    #print(astring)
                    astring = astring.replace(' ', '')
                    astring = astring.replace('\n', '')
                    astring = astring.replace('№', '')
                    astring = astring.replace(',', '')
                    astring = astring.replace('.', '')
                    #print(astring)
                    #print(astring)
                    query = query.replace(' ', '')
                    query = query.replace('\n', '')
                    query = query.replace('№', '')
                    query = query.replace(',', '')
                    query = query.replace('.', '')
                    if astring.lower() == query.lower():
                        captcha = get_captcha(s, captchakey)
                        #url = 'https://lk.rosreestr.ru/account-back/dictionary/OBJECT_TYPE_CODES?sortKey=code'
                        #res = s.get(url, verify=False)
                        #print(res)
                        url = 'https://lk.rosreestr.ru/account-back/on'
                        headers = {'Content-type': 'application/json'}
                        data = {
                            "filterType": "cadastral",
                            "cadNumbers": [i['cadnum']],
                            "captcha": f'{captcha}'
                        }
                        #print(json.dumps(data))
                        res = s.post(url, data=json.dumps(data), verify=False, headers=headers)
                        print(res)
                        result = json.loads(res.text)
                        #json.dumps(result['elements'])
                        try:
                            worksheet[f'A{key}']=result['elements'][0]['address']['region']
                            worksheet[f'B{key}']=0#result['elements'][0]['address']['dictrict']
                            worksheet[f'C{key}']=result['elements'][0]['address']['city']
                            worksheet[f'D{key}']=result['elements'][0]['address']['streetType'] + ' ' + result['elements'][0]['address']['street']
                            worksheet[f'E{key}']=result['elements'][0]['address']['house']
                            worksheet[f'F{key}']=result['elements'][0]['address']['apartment']
                            worksheet[f'G{key}']=result['elements'][0]['levelFloor']
                            worksheet[f'H{key}']=result['elements'][0]['cadNumber']
                            regDate = str(result['elements'][0]['regDate'])
                            #print(regDate[10:])
                            intDate = int(regDate[:-3])
                            worksheet[f'I{key}']=datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y')
                            worksheet[f'J{key}']=result['elements'][0]['cadCost']
                            regDate = str(result['elements'][0]['cadCostDeterminationDate'])
                            #print(regDate[10:])
                            intDate = int(regDate[:-3])
                            worksheet[f'K{key}']=datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y')
                            regDate = str(result['elements'][0]['cadCostRegistrationDate'])
                            #print(regDate[10:])
                            intDate = int(regDate[:-3])
                            worksheet[f'L{key}']=datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y')
                            temp = ''
                            intDate = ''
                            e = 0
                            try:
                                if result['elements'][0]['rights'][e] == 'null':
                                    break
                                regDate = str(result['elements'][0]['rights'][e]['rightRegDate'])
                                #print(regDate[10:])
                                intDate = int(regDate[:-3])
                                temp = temp + result['elements'][0]['rights'][e]['rightTypeDesc'] + " № " + result['elements'][0]['rights'][e]['rightNumber'] + " от " + datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y') + '\n'
                                e = e + 1
                            except IndexError:
                                worksheet[f'M{key}']=temp
                            finally:
                                worksheet[f'M{key}']=temp
                                worksheet[f'N{key}']=datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y')
                            temp = ''
                            e = 0
                            try:
                                if result['elements'][0]['encumbrances'][e] == 'null':
                                    break
                                regDate = str(result['elements'][0]['encumbrances'][e]['startDate'])
                                #print(regDate[10:])
                                intDate = int(regDate[:-3])
                                temp = temp + result['elements'][0]['encumbrances'][e]['typeDesc'] + " № " + result['elements'][0]['encumbrances'][e]['encumbranceNumber'] + " от " + datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y') + '\n'
                                e = e + 1
                            except IndexError:
                                worksheet[f'O{key}']=temp
                            finally:
                                worksheet[f'O{key}']=temp
                                worksheet[f'P{key}']=datetime.utcfromtimestamp(intDate).strftime('%d.%m.%Y')
                            key = key + 1
                            break
                        except:
                            continue
                        finally:
                            wb.save('test.xlsx')
                    else:
                        continue
        else:
            continue