import os
import xlwt
from xml.dom import minidom


def init_xls(sheetname):
    font = xlwt.Font()
    font.name = 'SimSun'
    style = xlwt.XFStyle()
    style.font = font

    fxls = xlwt.Workbook(encoding='utf-8')
    table = fxls.add_sheet(sheetname, cell_overwrite_ok=True)

    return fxls, table


def read_string_xml(xmlname):
    sdict = {}
    doc = minidom.parse(xmlname)
    stringtaglen = len(doc.getElementsByTagName('string'))
    itemlen = len(doc.getElementsByTagName('item'))

    for idx in range(0, stringtaglen):
        nameattr = doc.getElementsByTagName('string')[idx].getAttribute('name')
        fc = doc.getElementsByTagName('string')[idx].firstChild
        if fc is None:
            content = ""
        else:
            try:
                content = fc.data
            except AttributeError:
                content = ""
                for d in doc.getElementsByTagName('string')[idx].childNodes:
                    if d.firstChild is None:
                        content += d.data
                    else:
                        content += d.firstChild.data
        sdict[nameattr] = content

    for idx in range(0, itemlen):
        nameattr = doc.getElementsByTagName('item')[idx].getAttribute('name')
        fc = doc.getElementsByTagName('item')[idx].firstChild
        if fc is None:
            content = ""
        else:
            try:
                content = fc.data
            except AttributeError:
                content = ""
                for d in doc.getElementsByTagName('item')[idx].childNodes:
                    if d.firstChild is None:
                        content += d.data
                    else:
                        content += d.firstChild.data
        sdict[nameattr] = content
    return sdict


def main():
    inputfilepath = r'C:\Users\siwei.yan\Desktop\res'
    xls, table = init_xls('Language')
    table.write(0, 0, 'ResourceName')
    dir_list = os.listdir(inputfilepath)
    list = []
    dict = {}
    i = 1
    for file in dir_list:
        if file.__contains__('values'):
            filename = inputfilepath + '/' + file + '/' + 'strings.xml'
            if not os.path.exists(filename):
                continue
            resourcedict = read_string_xml(filename)
            for resource in resourcedict:
                if resource not in list:
                    list.append(resource)
                    dict[resource] = i
                    table.write(i, 0, resource)
                    i = i + 1
    k = 1
    for file in dir_list:
        if file.__contains__('values'):
            filename = inputfilepath + '/' + file + '/' + 'strings.xml'
            if not os.path.exists(filename):
                continue
            resourcedict = read_string_xml(filename)
            colname = file[7:]
            if colname == "":
                colname = "default"
            table.write(0, k, colname)
            for resource in resourcedict:
                table.write(dict[resource], k, resourcedict[resource])
            k = k + 1
    outputfilepath = r'C:\Users\siwei.yan\Desktop\result.xls'
    if os.path.exists(outputfilepath):
        os.remove(outputfilepath)
    xls.save(outputfilepath)


if __name__ == '__main__':
    main()
