import binascii
import shutil
import requests
from selenium import webdriver
from bs4 import BeautifulSoup
from PIL import Image

print("Enter category number and page number:")

categories = {
    "Кофе": "kofe-9372",
    "Чай": "chay-9373",
    "Какао, горячий шоколад": "kakao-goryachiy-shokolad-9371",
    "Цикорий, кэроб": "tsikoriy-kerob-9374",
    "Кофейные напитки": "kofeynye-napitki-9493",
    "Фильтры для заваривания чая": "filtry-dlya-zavarivaniya-chaya-1588",
    "Чайные напитки": "chaynye-napitki-33244",
    "Шоколад и шоколадные батончики": "shokolad-i-shokoladnye-batonchiki-9393",
    "Конфеты": "konfety-30695",
    "Печенье, пряники и вафли": "pechene-pryaniki-i-vafli-9381",
    "Выпечка и сдоба": "vypechka-i-sdoba-9376",
    "Зефир, пастила": "zefir-pastila-9388",
    "Мармелад": "marmelad-9390",
    "Торты, пирожные, бисквиты, коржи": "torty-pirozhnye-biskvity-korzhi-9386",
    "Хлеб, лаваши, лепешки": "hleb-lavashi-lepeshki-9377",
    "Восточные сладости, нуга": "vostochnye-sladosti-nuga-9379",
    "Шоколадная и ореховая пасты": "shokoladnaya-i-orehovaya-pasty-9397",
    "Леденцы": "ledentsy-9389",
    "Жевательная резинка": "zhevatelnaya-rezinka-9380",
}

index = 1

for category_name in categories:
    print('[%s] ' % index + category_name)
    index += 1

category_index = int(input()) - 1
page_number = int(input())
category_name = list(categories.keys())[category_index]
category_link = list(categories.values())[category_index]

print('Start writing...')

page_title = ''
name = ''
article = ''
image = ''
price = ''
report_begin = "ОТЧЕТ О ПРОДУКЦИИ ИНТЕРНЕТ МАГАЗИНА ОЗОН ПО ВЫБОРКЕ «ПРОДУКТЫ ПИТАНИЯ»"
report_end = "КОНЕЦ ДОКУМЕНТА"
domain = 'https://www.ozon.ru'
url = domain + '/category/' + category_link + '/?page=' + str(page_number)

driver = webdriver.Chrome()
driver.get(url)
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

page_title = soup.find('title').text
div = soup.find('div', {'data-widget': 'searchResultsV2'})
products_list = list()
links_list = list()

if div:
    div = div.find('div')
    if div:
        products_list = div.findChildren('div', recursive=False)

filename = 'products.rtf'
out_file = open(filename, 'w')
out_text = """{\\rtf1\\ansi\\deff0
%s\\line\\par
%s %s %s\\line\\par
\\trowd
\\cellx1000
\\cellx2000
\\cellx3000
\\cellx4000
Наименование\\intbl\\cell
Артикул\\intbl\\cell
Картинка\\intbl\\cell
Цена в руб\\intbl\\cell
\\row
""" % (report_begin, page_title, page_number, category_name)

for tag in products_list:
    div = tag.find('div')
    if div:
        a = div.find('a')
        if a:
            links_list.append(a['href'])

for link in links_list:
    driver.get(domain + link)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    h1 = soup.find('h1')
    if h1:
        name = h1.text
    div = soup.find('div', {'data-widget': 'webCharacteristics'})
    if div:
        div = div.find('div')
        if div:
            div = div.find('div')
            if div:
                dls = div.findChildren('div')
                if dls and len(dls) > 1:
                    for dl in dls[1]:
                        if dl:
                            dt = dl.find('dt')
                            if dt:
                                span = dt.find('span')
                                if span and span.text == 'Артикул':
                                    dd = dl.find('dd')
                                    if dd:
                                        article = dd.text
    div = soup.find('div', class_='magnifier-container')
    if div:
        img = div.find('img')
        if img:
            image = img['src']
    div = soup.find('div', {'data-widget': 'webSale'})
    if div:
        div = div.find('div')
        if div:
            div = div.find('div')
            if div:
                div = div.find('div')
                if div:
                    div = div.find('div')
                    if div:
                        div = div.find('div')
                        if div:
                            span = div.find('span')
                            if span:
                                price = span.text.split('\xa0₽')[0]
    r = requests.get(image, stream=True)
    filename = 'file.jpg'
    if r.status_code == 200:
        with open(filename, 'wb') as f:
            r.raw.decode_content = True
            shutil.copyfileobj(r.raw, f)
    f.close()
    new_image = Image.open(filename)
    new_image = new_image.resize((100, 100))
    new_image.save(filename)
    im = Image.open(filename)
    width, height = im.size
    im.close()
    with open(filename, 'rb') as f:
        content = f.read()
        image = binascii.hexlify(content)
        f.close()
    image = image.decode("utf-8")
    out_text += """
    \\trowd
    \\cellx1000
    \\cellx2000
    \\cellx3000
    \\cellx4000
    %s\\intbl\\cell
    %s\\intbl\\cell
    {\\pict\\picw %s\\pich %s\\picwgoal 100\\pichgoal 100\\pngblip %s}\\intbl\\cell
    %s\\intbl\\cell
    \\row
    """ % (name, article, width, height, image, price)
    print("Write: %s" % name)

out_text += """\\pard\\par
Отчет страницы %s\\line\\par%s}
""" % (len(links_list), report_end)

out_file.write(out_text)
out_file.close()
driver.close()