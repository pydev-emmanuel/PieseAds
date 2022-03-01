import cloudscraper
import requests
from bs4 import BeautifulSoup as bs


def get_product_url(product_code):
    scraper = cloudscraper.create_scraper()
    while True:
        url = scraper.get(f"https://www.autodoc24.ro/search?keyword={product_code}")
        if url.status_code == 200:
            break
        else:
            pass
    while True:
        html_code = bs(url.content, "lxml")
        tag = html_code.find("div", class_="name").a
        product_link = tag["href"]
        if product_link is None:
            pass
        else:
            break
    return product_link


def get_product_information(product_link):
    scraper = cloudscraper.create_scraper()
    while True:
        url = scraper.get(product_link)
        if url.status_code == 200:
            break
        else:
            pass
    html_code = bs(url.content, "lxml")
    product_information = {
        "Title": None,
        "Description": None,
        "Price": None
    }

    beta_title = html_code.find("div", class_="product-block__description__title product-block__equal-height-wrap").h2.span.text
    title = beta_title.replace("Arc spiral", "")
    print(title)
    product_information["Title"] = f"Arc spiral suspensie {title}"
    mini_description = html_code.find("span", class_="product-block__description__title-small").text
    for char in mini_description:
        if char == " ":
            print(mini_description.index(char))
        else:
            break
    print(mini_description)
    price = html_code.find("p", class_="product-new-price").text

    description =\
        f"""<h4>{title}</h4>
            <h3>
    
    """














print(get_product_information(get_product_url("4082937")))