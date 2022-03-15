import os
import xlsxwriter
from bs4 import BeautifulSoup as soup


directory = "C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS"

# for html_file in os.listdir(directory):
#     # print(html_file)


def product_data(html_file):
    oem_equivalent = []
    html = open(html_file, "r")
    contents = html.read()
    bs_content = soup(contents, "lxml")
    product_details = {
    }
    pozitie_montare = None
    pozitie_fixare = None
    lungime = None
    grosime = None
    diametru = None
    greutate = None
    intercar_price = bs_content.find_all("div", class_="quantity__amount")[1].text
    try:
        price = float(intercar_price) - float(intercar_price)*5/100
    except ValueError:
        price = intercar_price
    price = int(price)
    for tag in bs_content.find_all("td", class_="datatable__item datatable__item--wrapped"):
        if "Front" in tag.text:
            pozitie_montare = "FATA"
        elif "Rear" in tag.text:
            pozitie_montare = "SPATE"
    fixare = bs_content.find_all("a", class_="productfeatures__linkfeature")
    try:
        pozitie_check = f"{fixare[1].text} {fixare[2].text}"
        if "eft" and "ight" in pozitie_check:
            pozitie_fixare = "STANGA / DREAPTA"
        elif "eft" in pozitie_check:
            pozitie_fixare = "STANGA"
        elif "ight" in pozitie_check:
            pozitie_fixare = "DREAPTA"
    except IndexError:
        pozitie_fixare = None
    except TypeError:
        pozitie_fixare = None
    try:
        for tag in bs_content.find_all(class_="datatable__item"):
            if "Length" in tag.text:
                lungime = tag.findNext("td").a.text
                lungime = lungime.replace("\n", "")
                lungime = lungime.strip()
        for tag in bs_content.find_all(class_="datatable__item"):
            if "Thickness" in tag:
                grosime = tag.findNext("td").a.text
                grosime = grosime.replace("\n", "")
                grosime = grosime.strip()
        for tag in bs_content.find_all(class_="datatable__item"):
            if "Outer diameter" in tag:
                diametru = tag.findNext("td").a.text
                diametru = diametru.replace("\n", "")
                diametru = diametru.strip()
        for tag in bs_content.find_all(class_="datatable__item"):
            if "Weight" in tag:
                greutate = tag.findNext("td").span.span.text
    except AttributeError:
        pass
    for tag in bs_content.find_all("div", class_="refnumbers__listheader"):
        if "OEM part number equivalent" in tag.text:
            check = tag.findNext("div").ul
    try:
        for oem_tag in check.find_all("li", class_="refnumbers__item"):
            oem_equivalent.append(f"{oem_tag.find('span', class_='refnumbers__manufacturer').text} - {oem_tag.find('span', class_='refnumbers__refnumber').text}")
    except UnboundLocalError:
        oem_equivalent = []
    except AttributeError:
        pass
    product_details["pret"] = price
    product_details["pozitie_montare"] = pozitie_montare
    product_details["pozitie_fixare"] = pozitie_fixare
    product_details["lungime"] = lungime
    product_details["grosime"] = grosime
    product_details["diametru"] = diametru
    product_details["greutate"] = f"{greutate} kg"
    product_details["oem_equivalent"] = oem_equivalent
    return product_details


def product_aplicatii(html_file):
    html = open(html_file, "r")
    contents = html.read()
    bs_content = soup(contents, "lxml")
    aplicatii = {
    }
    aplicatii_list = []
    motor = None
    cod_motor = None
    ani_productie = None
    kW = None
    hp = None
    ccm = None
    for car_brand_tag in bs_content.find_all("div", class_="tree__branch js-tree-trigger is-open")[0]:
        car_brand = car_brand_tag.text
        next_tag = car_brand_tag.findNext("div")
        for leaf in next_tag.find_all("div", class_="tree__leaf js-tree-trigger is-open"):
            try:
                if leaf.findNext("div").ul["class"] == ['tree__list']:
                    pass
            except:
                name = leaf.text
                tabel = leaf.findNext("div")
                for masina in tabel.find_all("tr", class_="datatable__rowtd datatable__clickable js-clickable-row"):
                    for info in masina.find_all(class_="datatable__item"):
                        try:
                            if info.span.text == "Engine":
                                list_engine = list(info.text)
                                del list_engine[0:30]
                                engine = "".join(list_engine)
                                motor = engine.strip()
                            elif info.span.text == "Engine codes":
                                list_engine_code = list(info.text)
                                del list_engine_code[0:30]
                                engine_code = "".join(list_engine_code)
                                cod_motor = engine_code.strip()
                            elif info.span.text == "Production years":
                                production_years_list = list(info.text)
                                del production_years_list[0:30]
                                production_years = "".join(production_years_list)
                                production_years = production_years.replace("\n", "")
                                production_years = production_years.replace(" ", "")
                                ani_productie = production_years.strip()
                            elif info.span.text == "kW":
                                kw_list = list(info.text)
                                del kw_list[0:30]
                                kw = "".join(kw_list)
                                kW = f"{kw.strip()}kW"
                            elif info.span.text == "hp":
                                hp_list = list(info.text)
                                del hp_list[0:30]
                                horse = "".join(hp_list)
                                hp = f"{horse.strip()}cp"
                            elif info.span.text == "ccm":
                                ccm_list = list(info.text)
                                del ccm_list[0:30]
                                centimetricub = "".join(ccm_list)
                                ccm = f"{centimetricub.strip()}"
                        except:
                            pass
                    aplicatii_list.append([motor, cod_motor, ani_productie, kW, hp])
                aplicatii[f"{car_brand} {name}"] = aplicatii_list
                aplicatii_list = []
    return aplicatii


def descriere(aplicatii, product_details, cod_produs):
    pozitie_montare = product_details["pozitie_montare"]
    pozitie_fixare = product_details["pozitie_fixare"]
    lungime = product_details["lungime"]
    grosime = product_details["grosime"]
    diametru = product_details["diametru"]
    greutate = product_details["greutate"]
    oem_equivalent = product_details["oem_equivalent"]
    tabel_compatibilitate = []
    tabel_echivalente = []
    for key, value in aplicatii.items():
        for val in value:
            val_string = " ".join(val)
            tabel_compatibilitate.append(f"<div><b>{key} {val_string}</b></div>")
    tabel_compatibilitate = " ".join(tabel_compatibilitate)
    for ech in oem_equivalent:
        ech_string = " ".join(ech)
        tabel_echivalente.append(f"<div><b>{ech_string}</b></div>")
    tabel_echivalente = " ".join(tabel_echivalente)
    descriere = f"""<h2>Arc suspensie LESJOFORS {cod_produs}</h2><br>
        <div><br></div>
        <div><b>Se pot trimite poze cu piesa la cerere</b></div>
        <div><b>Piesa in stoc</b></div>
        <h3><u>Informatii produs:</u></h3>
        <div><b>Pozitie montare: {pozitie_montare}</b></div>
        <div><b>Pozitie fixare: {pozitie_fixare}</b></div>
        <div><b>Lungime: {lungime}</b></div>
        <div><b>Grosime: {grosime}</b></div>
        <div><b>Diametru exterior: {diametru}</b></div>
        <div><b>Greutate: {greutate}</b></div>
        <div><br></div>
        <h3><u>Masini compatibile:</u></h3>
        {tabel_compatibilitate}
        <div><br></div>
        <div><br></div>
        <h3>Echivalente coduri original:</h3>
        {tabel_echivalente}
        """
    return descriere


workbook = xlsxwriter.Workbook("C:\\Users\\Gh0sT\\Desktop\\WORKBOOK\\LESJOFORS.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "TITLU")
worksheet.write(0, 1, "CATEGORIE")
worksheet.write(0, 2, "DESCRIERE")
worksheet.write(0, 3, "MONEDA")
worksheet.write(0, 4, "PRET")
worksheet.write(0, 5, "CANTITATE")
worksheet.write(0, 6, "URL_POZA")
workbook.close()
directory = "C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS"
row = 1
for html_file in os.listdir(directory):
    print(row)
    cod_produs = html_file.replace(".html", "")
    print(cod_produs)
    html = f"{directory}\\{html_file}"
    product_details = product_data(html)
    aplicatii_produs = product_aplicatii(html)
    descriere_anunt = descriere(aplicatii_produs, product_details, cod_produs)
    pret = product_details["pret"]
    pozitie_montare = product_details["pozitie_montare"]
    pozitie_fixare = product_details["pozitie_fixare"]
    lungime = product_details["lungime"]
    grosime = product_details["grosime"]
    diametru = product_details["diametru"]
    greutate = product_details["greutate"]
    oem_equivalent = product_details["oem_equivalent"]
    for key, value in aplicatii_produs.items():
        for val in value:
            titlu = f"Arc suspensie {pozitie_montare} {pozitie_fixare} - {key} {' '.join(val)}  LESJOFORS {cod_produs}"
            worksheet.write(row, 0, titlu)
            worksheet.write(row, 1, "Arc spiral")
            worksheet.write(row, 2, descriere_anunt)
            worksheet.write(row, 3, "RON")
            worksheet.write(row, 4, pret)
            worksheet.write(row, 5, "1")
            worksheet.write(row, 6, "https://ibb.co/BLQc2ky")
            row += 1



