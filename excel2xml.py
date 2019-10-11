import os
import xlrd


def main():
    input_excel_path = r'C:\Users\siwei.yan\Desktop\test.xls'
    out_dir = r'C:\Users\siwei.yan\Desktop\res'
    if not os.path.exists(input_excel_path):
        return
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)
    excel = xlrd.open_workbook(input_excel_path)
    sheet = excel.sheet_by_index(0)
    language_count = sheet.ncols
    stringId_count = sheet.nrows
    language_index = 1
    while language_index < language_count:
        stringId_index = 1
        language_dir = 'values-' + sheet.cell(stringId_index - 1, language_index).value
        values_dir = out_dir + '\\' + language_dir
        if not os.path.exists(values_dir):
            os.mkdir(values_dir)
        xml_name = values_dir + '\\' + 'strings.xml'
        f = open(xml_name, 'w', encoding='utf8')
        content_head = '<?xml version="1.0" encoding="utf-8"?>' + '\n' + '<resources>' + '\n'
        f.write(content_head)
        while stringId_index < stringId_count:
            line_content_fixed = "<string name=\"" + sheet.cell(stringId_index, 0).value + "\">"
            line_content = line_content_fixed + sheet.cell(stringId_index, language_index).value + "</string>" + "\n"
            f.write(line_content)
            stringId_index = stringId_index + 1
        f.write("</resources>")
        language_index = language_index + 1


if __name__ == '__main__':
    main()

