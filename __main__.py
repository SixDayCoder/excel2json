import xlrd
from table_utils import *

def json_output(json_path, json_output_lst):
    if len(json_output_lst) == 1:
        return
    output = '[\n\t'
    for index, json_str in enumerate(json_output_lst):
        if index == len(json_output_lst) - 1:
            output = output + "{0}\n".format(json_str)
        else:
            output = output + "{0},\n\t".format(json_str)
    output = output + ']'
    with open(json_path, 'w+') as f:
        f.write(output)


def excel2json(excel_path, json_path):
    # open excel
    data = xlrd.open_workbook(excel_path)
    table = data.sheets()[0]
    meta_data = parse_excel_table_meta(table)
    # parse excel content expect excel head
    row = TABLE_CONTENT_START_INDEX
    json_output_lst = []
    while row < table.nrows:
        row_data = get_table_row_data(meta_data, table, row)
        json_str = parse_table_row(meta_data, table, row, row_data)
        json_output_lst.append(json_str)
        row += 1
    # output
    json_output(json_path, json_output_lst)


def main():
    excel_path = "Test.xlsx"
    excel2json(excel_path, "out.json")


if __name__ == '__main__':
    main()