import datetime


def get_xml_node(node, name):
    return node.getElementsByTagName(name) if node else []


def get_node_value(node, index=0):
    return node.childNodes[index].nodeValue if node else ''


def get_time(str):
    temp = str.rsplit("-", 1)
    res = temp[0].split("T")
    return res[0] + " " + res[1]


def add_fifteen_hours(str):
    dt = datetime.datetime.strptime(str, "%Y-%m-%d %H:%M:%S")
    out_date = (dt + datetime.timedelta(hours=15)).strftime("%Y-%m-%d")
    return out_date