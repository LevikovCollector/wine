import datetime
import pandas
import os
from collections import defaultdict
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

FOUNDATION_YEAR = 1920


def create_product_info() -> dict:
    """
    Function for read document and return dict with product information.
    :return: dict with product information
    """

    excel_data_df = pandas.read_excel(os.getenv("PATH_TO_DOCUMENT"),
                                      sheet_name=os.getenv("DOCUMENT_SHEET"),
                                      na_values=['nan'],
                                      keep_default_na=False)
    product_info = defaultdict(list)
    for category, name, kind, cost, img, stock in excel_data_df.values:
        product_info[category].append({'name': name,
                                       'kind': kind,
                                       'cost': cost,
                                       'img': img,
                                       'stock': stock})
    return product_info


def get_year_word(years: int) -> str:
    """
    Function generate year label

    :param years: it`s number of year we check
    :return: string with right label
    """

    last_two = years % 100
    last_one = years % 10

    if 11 <= last_two <= 19:
        return "лет"
    elif last_one == 1:
        return "год"
    elif 2 <= last_one <= 4:
        return "года"
    else:
        return "лет"


if __name__ == "__main__":
    load_dotenv()
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('wine_template.html')
    company_age = datetime.datetime.now().year - FOUNDATION_YEAR

    rendered_page = template.render(
        company_age=f"{company_age} {get_year_word(company_age)}",
        product_info=create_product_info()
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()