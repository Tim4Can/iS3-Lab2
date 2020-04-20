from win32com.client import Dispatch
import os
import csv


def batch_doc_to_docx(files):
    files_transformed = set()

    try:
        word = Dispatch("Kwps.Application")
    except:
        raise ValueError("此电脑上未安装WPS！")

    for file in files:
        try:
            doc = word.Documents.Open(file)
            new_file_name = "{}x".format(file)
            doc.SaveAs(new_file_name, 12)
            files_transformed.add(new_file_name)
            doc.Close()
        except IOError:
            print("读取文件异常：" + file)
    # word.Quit()
    return files_transformed


def check_output_file(output_path, header):
    if not os.path.exists(output_path):
        with open(output_path, "w", encoding="utf_8_sig", newline="") as f:
            w = csv.DictWriter(f, header)
            w.writeheader()


def parse_prefix(name):
    prefix = None
    if "泸水" in name:
        prefix = "L"
    elif "保山" in name:
        prefix = "B"

    if "左" in name:
        prefix = "Z"
    elif "右" in name:
        prefix = "Y"
    elif "斜" in name and prefix is not None:
        prefix += "X"
    return prefix


def parse_GSI_CHAI_and_GSI_INTE(name, GSI_INTE):
    prefix = parse_prefix(name)
    split = "～"
    GSI_INTE = GSI_INTE.split("（")[0].split(split)
    GSI_CHAI = GSI_INTE[0]

    GSI_CHAI = prefix + GSI_CHAI
    GSI_INTE = [prefix + part for part in GSI_INTE]
    GSI_INTE = split.join(GSI_INTE)
    return GSI_CHAI, GSI_INTE


def parse_SKTH_CHAI_and_SKTH_INTE(name, SKTH_INTE):
    prefix = parse_prefix(name)
    split = "～"
    SKTH_INTE = SKTH_INTE.split("（")[0].split(split)
    SKTH_CHAI = SKTH_INTE[0]

    SKTH_CHAI = prefix + SKTH_CHAI
    SKTH_INTE = [prefix + part for part in SKTH_INTE]
    SKTH_INTE = split.join(SKTH_INTE)
    return SKTH_CHAI, SKTH_INTE


def parse_GPRF_INTE(name, GPRF_INTE):
    prefix = parse_prefix(name)
    split = "～"
    GPRF_INTE = GPRF_INTE.split(split)
    GPRF_INTE = [prefix + part for part in GPRF_INTE]
    GPRF_INTE = split.join(GPRF_INTE)
    return GPRF_INTE
