import os
import sys
from re import sub
import jieba
import xlwings as xw
from translate import Translator
import re

#
translator = Translator(from_lang="Chinese", to_lang="English")
FIELD_STR = "//{cnFieldName}\nprivate String {fieldName};"
jieba.load_userdict("user_dict.txt")
SQL_STR = "`{fieldName}` VARCHAR ( 255 ) COMMENT '{cnFieldName}',"


def name_convert_to_camel(name: str) -> str:
    """下划线转驼峰(小驼峰)"""
    return re.sub(r'(_[a-z])', lambda x: x.group(1)[1].upper(), name)


def name_convert_to_snake(name: str) -> str:
    """驼峰转下划线"""
    if '_' not in name:
        name = re.sub(r'([a-z])([A-Z])', r'\1_\2', name)
    else:
        raise ValueError(f'{name}字符中包含下划线，无法转换')
    return name.lower()


def name_convert(name: str) -> str:
    """驼峰式命名和下划线式命名互转"""
    is_camel_name = True  # 是否为驼峰式命名
    if '_' in name and re.match(r'[a-zA-Z_]+$', name):
        is_camel_name = False
    elif re.match(r'[a-zA-Z]+$', name) is None:
        raise ValueError(f'Value of "name" is invalid: {name}')
    return name_convert_to_snake(name) if is_camel_name else name_convert_to_camel(name)


def toEnglish(strList):
    arr = []
    for str in strList:
        arr.append(translator.translate(str))
    return "_".join(arr)


def split(str):
    return jieba.cut(str)


def camelCase(string):
    if not string:
        return ''
    if '' == string:
        return ''
    string = sub(r"(_|-)+", " ", string).title().replace(" ", "").replace(".", "").replace(":", "")
    return string[0].lower() + string[1:]


def readExcelTitle(fileUrl):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(fileUrl)
    sheet = wb.sheets[0]
    # 读取第一行数据
    rng = sheet.range('a1').expand('table')
    ncols = rng.columns.count
    fst_col = sheet[0, :ncols].value
    wb.close()
    app.kill()

    return fst_col


def ctrlStr(cnName, fieldStr, isXh):
    try:
        nameList = split(cnName)
        str = toEnglish(nameList)
        fieldName = camelCase(str)
        if isXh:
            fieldName = name_convert_to_snake(fieldName)
        fieldStr = fieldStr.format(cnFieldName=cnName, fieldName=fieldName)
        return fieldStr
    except Exception as e:
        print(cnName + "失败！！！！！！！")


if __name__ == '__main__':

    if len(sys.argv) == 1:
        print("请带上excel参数执行！")
        sys.exit(1)

    print('excel文件路径 :%s' % sys.argv[1])
    fileUrl = sys.argv[1]
    base_name = os.path.basename(fileUrl)
    fileName = base_name.split(".")[0]

    try:
        filedNameList = readExcelTitle(fileUrl)
    except Exception as e:
        print("请安装excel或wps！")
        sys.exit(1)

    # 生成java类
    str = ''
    for cnName in filedNameList:
        str = str + ctrlStr(cnName, FIELD_STR, False) + "\n"

    with   open('javafile', encoding='UTF-8') as f:
        content = f.read()
        con = content.format(allFiled=str, CN_TABLENAME=fileName)
        with  open(fileName + ".java", encoding='UTF-8', mode="w") as fileobj:
            fileobj.write(con)

    # 生成 create sql
    str = ''
    for cnName in filedNameList:
        str = str + ctrlStr(cnName, SQL_STR, True) + "\n"
    with open('sqlfile', encoding='UTF-8') as f:
        content = f.read()
        con = content.format(allFiled=str, CN_TABLENAME=fileName)
        with  open(fileName + ".sql", encoding='UTF-8', mode="w") as fileobj:
            fileobj.write(con)
