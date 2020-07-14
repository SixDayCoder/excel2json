# coding=utf-8

TABLE_CONTENT_START_INDEX = 3

class TableFieldMeta:

    def __init__(self):
        self.key_lst = []
        self.type_lst = []
        self.array_dict = {}


def format_table_field_data(v, t):
    if t == 'INT' or t == 'INT64':
        v = int(v)
    if t == 'FLOAT':
        v = float(v)
    if t == 'STRING':
        v = "\"" + v.encode('utf-8') + "\""
    if t == 'DATE':
        v = "\"" + v.encode('utf-8') + "\""
    return v

def is_array_field(k):
    index = k.find('_')
    return index > 0

def get_real_key_name(k):
    index = k.find('_')
    if index > 0:
        k = k[0:index]
    return k

def remove_last_comma(s):
    index = s.rfind(',')
    if index > 0:
        tmp = s[0:index]
        tmp = tmp + s[index + 1:]
        return tmp
    return s

def parse_excel_table_meta(excel_table):
    cell_field_name_lst = excel_table.row(0)
    cell_field_type_lst = excel_table.row(1)

    # parse key
    field_key_lst = []
    for i, name in enumerate(cell_field_name_lst):
        field_key = excel_table.cell_value(0, i).encode('utf-8')
        field_key_lst.append(field_key)

    # parse type
    field_type_lst = []
    for i, t in enumerate(cell_field_type_lst):
        field_type = excel_table.cell_value(1, i).encode('utf-8')
        field_type_lst.append(field_type)

    # parse array element
    field_array_dict = {}
    for i, k in enumerate(field_key_lst):
        if is_array_field(k):
            k = get_real_key_name(k)
            if field_array_dict.get(k):
                field_array_dict[k]["index"].append(i)
            else:
                field_array_dict[k] = {}
                field_array_dict[k]["type"] = field_type_lst[i]
                field_array_dict[k]["index"] = []
                field_array_dict[k]["index"].append(i)

    # fill table meta data
    data = TableFieldMeta()
    data.key_lst = field_key_lst
    data.type_lst = field_type_lst
    data.array_dict = field_array_dict
    return data


def parse_table_row(meta_data, table, row, row_data):
    json_str = "{\n"
    # parse no array data
    for index, k in enumerate(meta_data.key_lst):
        t = meta_data.type_lst[index]
        v = row_data[index]
        # skip array data
        if is_array_field(k):
            continue
        json_str = json_str + "\t\t\"{0}\" : {1},\n".format(k, v)

    # parse array data
    have_parsed_array_key = {}
    for index, k in enumerate(meta_data.key_lst):
        if not is_array_field(k):
            continue
        k = get_real_key_name(k)
        # skip parsed key
        if have_parsed_array_key.get(k):
            continue
        array_str = "["
        for col in meta_data.array_dict[k]["index"]:
            v = table.cell_value(row, col)
            v = format_table_field_data(v, meta_data.array_dict[k]["type"])
            array_str = array_str + "{0},".format(v)
        array_str  = array_str + "]"
        array_str = remove_last_comma(array_str)
        json_str = json_str + "\t\t\"{0}\" : {1},\n".format(k, array_str)
        have_parsed_array_key[k] = True

    # result
    json_str = remove_last_comma(json_str)
    json_str = json_str + "\t}"
    return json_str


def get_table_row_data(meta_data, table, row):
    field_val_lst = []
    index = 0
    while index < len(meta_data.type_lst):
        t = meta_data.type_lst[index]
        v = table.cell_value(row, index)
        field_val_lst.append(format_table_field_data(v, t))
        index += 1
    return field_val_lst