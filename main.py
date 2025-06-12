import datetime
import pandas
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

DATE_OF_FOUNDATION = 1920


def create_product_dict() -> dict:
    """
    Function for read document and return dict with product information.
    :return: dict with product information
    """

    excel_data_df = pandas.read_excel('wine3.xlsx', sheet_name='Лист1')
    product_dict = defaultdict(list)
    for data in excel_data_df.values:
        category = data[0]
        name = check_value(data[1])
        kind = check_value(data[2])
        cost = check_value(data[3])
        img = check_value(data[4])
        stock = check_value(data[5])

        product_dict[category].append({'name': name,
                                       'kind': kind,
                                       'cost': cost,
                                       'img': img,
                                       'stock': stock})
    return product_dict


def check_value(value):
    """
    Function check value and return empty string if value nan, or return value

    :param value: - value for check
    :return: return empty string or value
    """
    if str(value) == 'nan':
        return ""
    return value


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


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('wine_template.html')
company_age = datetime.datetime.now().year - DATE_OF_FOUNDATION

rendered_page = template.render(
    company_age=f"{company_age} {get_year_word(company_age)}",
    product_dict=create_product_dict()
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)


server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()

