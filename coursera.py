import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from openpyxl.compat import range
from lxml import html
import sys


def get_courses_list():

    try:
        coursera_xml = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    except requests.exceptions.RequestException:
        coursera_xml = None

    tree = html.fromstring(coursera_xml.content)
    courses_list =  tree.xpath('//loc/text()')

    return courses_list[:20]

def get_course_info(course_slug):

    try:
        page = requests.get(course_slug)
        page.encoding = 'UTF-8'
    except requests.exceptions.RequestException:
        page = None

    soup = BeautifulSoup(page.text, 'html.parser')
    course_title = soup.find('title').get_text().split(' |')[0]
    course_language = soup.find('div', class_='rc-Language').get_text().split(',')[0]
    course_started = soup.find('div', class_='startdate rc-StartDateString caption-text').get_text()
    course_commitment_raw = soup.find_all('div', class_='rc-BasicInfo')[0].get_text()
    course_commitment = re.findall(r'Commitment(.*)Language',str(course_commitment_raw))

    if len(course_commitment) == 0:
        course_commitment = 'None'
    else:
        course_commitment = course_commitment[0]

    course_mark = re.findall(r'"averageFiveStarRating":(\S{3}),' , str(soup))

    if len(course_mark) == 0:
        course_mark = 'None'
    else:
        course_mark = course_mark[0]

    course_info = [course_title,course_language,course_started,course_commitment,course_mark,course_slug]

    return course_info


def output_courses_info_to_xlsx(filepath):

    courses_list = get_courses_list()

    courses = []

    for course_slug in courses_list:
        courses.append(get_course_info(course_slug))

    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Title'
    ws['B1'] = 'Language'
    ws['C1'] = 'Started'
    ws['D1'] = 'Commitment'
    ws['E1'] = 'Avg.mark'
    ws['F1'] = 'URL'

    for row in range(2, len(courses)+2):
        for col in range(1, 7):
            _ = ws.cell(column=col, row=row, value="{0}".format(courses[row-2][col-1]))

    wb.save(filepath)


if __name__ == '__main__':
    pass

filepath = 'c:/1/coursera.xlsx'
output_courses_info_to_xlsx(filepath)
