import requests
from xml.dom.minidom import parse
import time
import xlsxwriter
import settings
from utils import get_xml_node, get_node_value, get_time, add_fifteen_hours


class review:
    updated = ''
    rating = ''
    version = ''
    region = ''
    author = ''
    title = ''
    content = ''

    def __init__(self, updated, rating, version, region, author, title, content):
        self.updated = updated
        self.rating = rating
        self.version = version
        self.region = region
        self.author = author
        self.title = title
        self.content = content


def get_review_list():
    review_list = []
    # > pip install lxml -i http://pypi.douban.com/simple --trusted -host pypi.douban.com

    total = 176
    all_page = 6
    for n in range(1, total):
        for page in range(1, all_page):
            region = settings.region[n]
            region_name = settings.region_name[str(region)]
            url = 'https://itunes.apple.com/' + str(region) + '/rss/customerreviews/page=' + \
                  str(page) + '/id=------' + '/sortby=mostrecent/xml'
            print("start to crawl: " + url)

            response = requests.get(url, settings.headers)
            if response.status_code == 200:
                with open('VeSync.xml', 'wb') as fp:
                    fp.write(response.content)

            dom_tree = parse("VeSync.xml")
            root_node = dom_tree.documentElement
            entrys = root_node.getElementsByTagName("entry")
            if entrys:
                for entry in entrys:
                    updated = get_xml_node(entry, 'updated')
                    updated_str = get_node_value(updated[0])
                    temp = get_time(updated_str)
                    temp_str = add_fifteen_hours(temp)

                    title = get_xml_node(entry, 'title')
                    title_str = get_node_value(title[0])

                    content = get_xml_node(entry, 'content')
                    content_str = get_node_value(content[0])

                    rating = get_xml_node(entry, 'im:rating')
                    rating_str = get_node_value(rating[0])

                    version = get_xml_node(entry, 'im:version')
                    version_str = get_node_value(version[0])

                    author_name = ""
                    authors = entry.getElementsByTagName("author")
                    for author in authors:
                        name = get_xml_node(author, 'name')
                        name_str = get_node_value(name[0])
                        author_name = name_str

                    r = review(temp_str, rating_str, version_str, region_name, author_name, title_str, content_str)
                    review_list.append(r)
            else:
                break

    return review_list


def write_to_excel():

    time_title = time.strftime("%Y.%m.%d", time.localtime())
    workbook = xlsxwriter.Workbook('APP review-' + time_title + '.xlsx')
    worksheet = workbook.add_worksheet('ios')
    format = workbook.add_format()
    format.set_border(1)
    format_title = workbook.add_format()
    format_title.set_border(1)
    format_title.set_bg_color('#cccccc')
    format_title.set_align('left')
    format_title.set_bold()
    title = ['统计更新时间', '问题序号', '时间', '星级', '版本', '地区', '用户名', '标题', '内容原文']

    # 设置单元格宽度
    worksheet.set_column(0, 0, 12)
    worksheet.set_column(1, 1, 10)
    worksheet.set_column(2, 2, 12)
    worksheet.set_column(3, 3, 10)
    worksheet.set_column(4, 4, 10)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 30)
    worksheet.set_column(7, 7, 30)
    worksheet.set_column(8, 8, 100)
    worksheet.write_row('A1', title, format_title)

    no = 1
    total_count = 0
    review_list = get_review_list()
    review_list.sort(key=lambda x: x.updated, reverse=True)
    time_str = time.strftime("%Y-%m-%d", time.localtime())

    for obj in review_list:
        start_row = total_count + 1
        worksheet.write(start_row, 0, time_str, format)
        worksheet.write(start_row, 1, no, format)
        worksheet.write(start_row, 2, obj.updated, format)
        worksheet.write(start_row, 3, obj.rating, format)
        worksheet.write(start_row, 4, obj.version, format)
        worksheet.write(start_row, 5, obj.region, format)
        worksheet.write(start_row, 6, obj.author, format)
        worksheet.write(start_row, 7, obj.title, format)
        worksheet.write(start_row, 8, obj.content, format)
        total_count = total_count + 1
        no = no + 1

    workbook.close()
    print("totalCount: " + str(total_count))
    print("finish")

