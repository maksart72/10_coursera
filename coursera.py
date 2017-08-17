import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from lxml import html


def get_courses_list(courses_num):

    try:
        coursera_xml = requests.get(
            'https://www.coursera.org/sitemap~www~courses.xml')
    except requests.exceptions.RequestException:
        coursera_xml = None

    tree = html.fromstring(coursera_xml.content)
    courses_list = tree.xpath('//loc/text()')

    return courses_list[:courses_num]


def get_course_info(course_slug):

    try:
        page = requests.get(course_slug)
        page.encoding = 'UTF-8'
    except requests.exceptions.RequestException:
        page = None

    soup = BeautifulSoup(page.text, 'html.parser')
    course_title = soup.find('title').get_text().split(' |')[0]
    course_language = soup.find(
        'div', class_='rc-Language').get_text().split(',')[0]
    course_started = soup.find(
        'div', class_='startdate rc-StartDateString caption-text').get_text()
    course_commitment_raw = soup.find_all(
        'div', class_='rc-BasicInfo')[0].get_text()
    course_commitment = re.findall(
        r'Commitment(.*)Language', str(course_commitment_raw))

    if not course_commitment:
        course_commitment = 'None'
    else:
        course_commitment = course_commitment[0]

    course_mark = re.findall(r'"averageFiveStarRating":(\S{3}),', str(soup))

    if not course_mark:
        course_mark = 'None'
    else:
        course_mark = course_mark[0]

    course_info = [course_title, course_language, course_started,
                   course_commitment, course_mark, course_slug]

    return course_info


def output_courses_info_to_xlsx(filepath, courses_list):
    courses = []
    for course_slug in courses_list:
        courses.append(get_course_info(course_slug))

    courses_workbook = Workbook()
    courses_sheet = courses_workbook.active
    courses_sheet['A1'] = 'Title'
    courses_sheet['B1'] = 'Language'
    courses_sheet['C1'] = 'Started'
    courses_sheet['D1'] = 'Commitment'
    courses_sheet['E1'] = 'Avg.mark'
    courses_sheet['F1'] = 'URL'

    for row in range(2, len(courses) + 2):
        for col in range(1, 7):
            courses_sheet.cell(column=col, row=row, value="{0}".format(
                courses[row - 2][col - 1]))
    try:
        courses_workbook.save(filepath)
    except PermissionError:
        print('Can not write to file!')

if __name__ == '__main__':

    filepath = 'c:/1/coursera.xlsx'
    courses_num = 10
    courses_list = get_courses_list(courses_num)
    output_courses_info_to_xlsx(filepath, courses_list)
