import os
import xlsxwriter
from bs4 import BeautifulSoup as soup


directory = "C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS"

# for html_file in os.listdir(directory):
#     # print(html_file)


def retract_data(html_file):
    product_details = {
    }

    aplicatii = {
    }
    pret = None
    pozitie_montare = None
    pozitie_fixare = None
    lungime = None
    grosime = None
    diametru = None
    greutate = None
    oem_equivalent = []
    html = open(html_file, "r")
    contents = html.read()
    bs_content = soup(contents, "lxml")
    price = bs_content.find_all("div", class_="quantity__amount")[1].text
    for tag in bs_content.find_all("td", class_="datatable__item datatable__item--wrapped"):
        if "Front" in tag.text:
            pozitie_montare = "FATA"
        elif "Rear" in tag.text:
            pozitie_montare = "SPATE"
    fixare = bs_content.find_all("a", class_="productfeatures__linkfeature")
    pozitie_fixare = f"{fixare[1].text} {fixare[2].text}"
    if "eft" and "ight" in pozitie_fixare:
        pozitie_fixare = "STANGA / DREAPTA"
    elif "eft" in pozitie_montare:
        pozitie_fixare = "STANGA"
    elif "ight" in pozitie_montare:
        pozitie_fixare = "DREAPTA"
    for tag in bs_content.find_all(class_="datatable__item"):
        if "Length" in tag.text:
            lungime = tag.findNext("td").a.text
    for tag in bs_content.find_all(class_="datatable__item"):
        if "Thickness" in tag:
            grosime = tag.findNext("td").a.text
    for tag in bs_content.find_all(class_="datatable__item"):
        if "Outer diameter" in tag:
            diametru = tag.findNext("td").a.text
    for tag in bs_content.find_all(class_="datatable__item"):
        if "Weight" in tag:
            greutate = tag.findNext("td").span.span.text
    for tag in bs_content.find_all("div", class_="refnumbers__listheader"):
        if "OEM part number equivalent" in tag.text:
            check = tag.findNext("div").ul
    for oem_tag in check.find_all("li", class_="refnumbers__item"):
        oem_equivalent.append(f"{oem_tag.find('span', class_='refnumbers__manufacturer').text} - {oem_tag.find('span', class_='refnumbers__refnumber').text}")
    product_details["pret"] = price
    product_details["pozitie_montare"] = pozitie_montare
    product_details["pozitie_fixare"] = pozitie_fixare
    product_details["lungime"] = lungime
    product_details["grosime"] = grosime
    product_details["diametru"] = diametru
    product_details["greutate"] = greutate
    product_details["oem_equivalent"] = oem_equivalent










retract_data("C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS\\4095050.html")


# bs_content = soup("C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS\\4015660.html", "lxml")
# print(bs_content)
