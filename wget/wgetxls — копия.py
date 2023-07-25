import requests, os, sys, csv
from pathlib import Path

import xml.etree.ElementTree as ET
import time
import schedule
from openpyxl import Workbook




def GetBasketJob():

    dir_path = Path(os.path.dirname(os.path.realpath(__file__)))
    timestr = time.strftime("%Y%m%d-%H%M%S")
    reqfilepath = Path(dir_path, timestr+' Корзина.xls')
    revisedreqfile = reqfilepath.as_posix()
    
    cookies = {
        'BITRIX_SM_GUEST_ID': '314264',
        'BITRIX_SM_LAST_VISIT': '03.07.2023^%^2011^%^3A42^%^3A44',
        'BITRIX_SM_LAST_ADV': '5_Y',
        'BITRIX_SM_mybackurl': '^%^2F',
        'BITRIX_CONVERSION_CONTEXT_s1': '^%^7B^%^22ID^%^22^%^3A2^%^2C^%^22EXPIRE^%^22^%^3A1688417940^%^2C^%^22UNIQUE^%^22^%^3A^%^5B^%^22conversion_visit_day^%^22^%^5D^%^7D',
        'BITRIX_SM_GUEST_ID': '314264',
        'BITRIX_SM_LAST_VISIT': '03.07.2023^%^2011^%^3A42^%^3A42',
        'BITRIX_SM_LAST_ADV': '5_Y',
        'BITRIX_SM_mybackurl': '^%^2F',
        '_ym_uid': '1669026092136047434',
        '_ym_d': '1684823787',
        'BX_USER_ID': '949b16131b6dd3aea4654cad641d89b7',
        'BITRIX_SM_SALE_UID': '5223061',
        'BITRIX_SM_SALE_UID': '5223061',
        'BITRIX_SM_LOGIN': 'albertkovach',
        'BITRIX_SM_LOGIN': 'albertkovach',
        'BITRIX_SM_FLASH_MESSAGE': '1682414803',
        'PHPSESSID': 'n00XAZqT8RO0NGjnxa7rhjiia0IQqTZo',
        'BITRIX_SM_NCC': 'Y',
        'BITRIX_SM_NCC': 'Y',
        '_ym_isad': '2',
        'BITRIX_SM_SOUND_LOGIN_PLAYED': 'Y',
        'BITRIX_SM_SOUND_LOGIN_PLAYED': 'Y',
        '_ym_visorc': 'w',
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/114.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        # 'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://www.makel25.ru/bitrix/admin/sale_order.php?lang=ru',
        'Connection': 'keep-alive',
        # 'Cookie': 'BITRIX_SM_GUEST_ID=314264; BITRIX_SM_LAST_VISIT=03.07.2023^%^2011^%^3A42^%^3A44; BITRIX_SM_LAST_ADV=5_Y; BITRIX_SM_mybackurl=^%^2F; BITRIX_CONVERSION_CONTEXT_s1=^%^7B^%^22ID^%^22^%^3A2^%^2C^%^22EXPIRE^%^22^%^3A1688417940^%^2C^%^22UNIQUE^%^22^%^3A^%^5B^%^22conversion_visit_day^%^22^%^5D^%^7D; BITRIX_SM_GUEST_ID=314264; BITRIX_SM_LAST_VISIT=03.07.2023^%^2011^%^3A42^%^3A42; BITRIX_SM_LAST_ADV=5_Y; BITRIX_SM_mybackurl=^%^2F; _ym_uid=1669026092136047434; _ym_d=1684823787; BX_USER_ID=949b16131b6dd3aea4654cad641d89b7; BITRIX_SM_SALE_UID=5223061; BITRIX_SM_SALE_UID=5223061; BITRIX_SM_LOGIN=albertkovach; BITRIX_SM_LOGIN=albertkovach; BITRIX_SM_FLASH_MESSAGE=1682414803; PHPSESSID=n00XAZqT8RO0NGjnxa7rhjiia0IQqTZo; BITRIX_SM_NCC=Y; BITRIX_SM_NCC=Y; _ym_isad=2; BITRIX_SM_SOUND_LOGIN_PLAYED=Y; BITRIX_SM_SOUND_LOGIN_PLAYED=Y; _ym_visorc=w',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        # Requests doesn't support trailers
        # 'TE': 'trailers',
    }

    params = {
        'lang': 'ru',
        'mode': 'excel',
    }

    response = requests.get('https://www.makel25.ru/bitrix/admin/sale_basket.php', params=params, cookies=cookies, headers=headers)
    if response.status_code == 200:
        with open(revisedreqfile,'wb') as reqfile:
            reqfile.write(response.content)
            
        #messagebox.showinfo('Succes', 'File downloaded !')
        print(revisedreqfile)
    else:
        x=1
        #messagebox.showinfo('Ошибка', 'Запрос прервался с кодом ошибки: {0}'.format(response.status_code)) 
    




def Start():
    print("Started !")
    schedule.every(30).minutes.do(GetBasketJob)
    while True:
        schedule.run_pending()
        time.sleep(1)




Start()