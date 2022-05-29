import zipfile
import os
import stat
import shutil
from bs4 import BeautifulSoup
from constans import *


def file_read(file):
    try:
        with open(file, "rb") as FILE:
            data = FILE.read()
        return data
    except Exception:
        return "Error"


def file_write(file, data):
    try:
        with open(file, "wb") as FILE:
            FILE.write(data)
        return True
    except Exception:
        return "Error"


def find_tag(file, tag=TAG_PROTECTION_SHEET):
    soup = BeautifulSoup(file_read(file), "xml")
    find_tag = soup.find(tag)
    return find_tag


def remove_tag(file, tag=TAG_PROTECTION_SHEET):
    soup = BeautifulSoup(file_read(file), "xml")
    data = file_read(file)
    if data:
        search_tag = find_tag(file, tag)
        if search_tag:
            soup.select_one(tag).decompose()
            file_write(file, soup.prettify("utf-8"))
        else:
            return None
    return data


def get_info(file):
    data = file_read(file)
    soup = BeautifulSoup(data, "xml")

    find_all = soup.find_all("sheet")
    sheets = []
    for index, sheet in enumerate(find_all):
        number = index + 1

        name = sheet.get("name")
        state = sheet.get("state")
        if not state:
            state = "default"

        file_now = "{}\\sheet{}.xml".format(PATH_WORKSHEETS, number)
        protect = find_tag(file_now)

        sheets.append({"name": name,
                       "state": state,
                       "is_protect": "Yes" if protect else "No"})
    return sheets


def extract(file_name):
    ZIP = zipfile.ZipFile(file_name)
    ZIP.extractall(PATH_EXTRACT)
    ZIP.close()


def patching():
    count = 0
    get_info_sheets = get_info(PATH_WORKBOOK)
    for folder, _, files in os.walk(PATH_WORKSHEETS):
        if "_rels" not in folder:
            for file in files:
                sheets = get_info_sheets[count]
                path = folder + "\\" + file
                remove = remove_tag(path)
                name = repr(sheets["name"])
                if remove and remove != "Error":
                    print("\t[>] {} protection removed".format(name))
                elif remove == "Error":
                    print("\t[>] {} protection don't removed".format(name))
                else:

                    print("\t[>] {} protction not found".format(name))
                count += 1


def compress(file_name):
    new_zip = zipfile.ZipFile("{}_no_protect.xlsx".format(
                                    os.path.splitext(file_name)[0]), "w")

    for folder, _, files in os.walk(PATH_EXTRACT):
        for file in files:
            new_zip.write(os.path.join(folder, file),
                    os.path.relpath(os.path.join(folder, file), PATH_EXTRACT),
                    compress_type = zipfile.ZIP_DEFLATED)


def remove_extract(file_name):
    def remove_readonly(action, file_name, exception):
        os.chmod(file_name, stat.S_IWRITE)
        os.remove(file_name)
    shutil.rmtree(PATH_EXTRACT, onerror=remove_readonly)


def remove_hide():
    try:
        with open(PATH_WORKBOOK, "rb") as FILE:
            data = FILE.read()
        data_new = data.replace(b"state=\"hidden\"", b"")
        with open(PATH_WORKBOOK, "wb") as FILE:
            FILE.write(data_new)
        return True
    except Exception:
        return False
