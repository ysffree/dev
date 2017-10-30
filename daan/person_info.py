#!/usr/bin/env python3
# coding:utf-8
"""
usage:
        simple_info_bata.py 2017.02.23
        python3 simple_info_bata.py simple_id
        default:SERVER = "192.168.135.11"
                PORT = "8080"

        eg: python3 simple_info_bata.py S160000000
"""
import json
import time
import sys
import urllib.request
from collections import OrderedDict

SERVER = "app.smartquerier.com"
PORT = "8001"
def set_api():
    AF_API = "http://{SERVER}:{PORT}/bigdatams/rest/commonDB/list/sample/number=".format(SERVER=SERVER, PORT=PORT)
    MF_API = "http://{SERVER}:{PORT}/bigdatams/rest/commonDB/list/samplecollection/sampleNumber=".format(SERVER=SERVER, PORT=PORT)
    KF_API = "http://{SERVER}:{PORT}/bigdatams/rest/commonDB/list/patient/sampleNumber=".format(SERVER=SERVER, PORT=PORT)
    return AF_API, MF_API, KF_API

def url_open(simple_id = None, API = None):
    if simple_id == None:
        print("请输入查阅样本编号!")
        exit()
    url = API + simple_id
    try:
        response = urllib.request.urlopen(url)
    except:
        print('HTTP Status Not Found', file=sys.stderr)
    cag_data = response.read().decode('utf8')
    if not cag_data:
        print('Data Form Error')
    ret = json.loads(cag_data)
    if ret["errorMessage"] == "noDataError" or ret["status"] == "error" or ret["result"] == "":
        raise Exception("不存在的编号!", file=sys.stderr)
    return ret

def get_sample_type(typeOptionValue):  # 样本类型，1开头的为血液，2开头的为组织，3开头的为胸水
    if str(typeOptionValue).startswith('1'):
        return '血液'
    elif str(typeOptionValue).startswith('2'):
        return '组织'
    elif str(typeOptionValue).startswith('3'):
        return '胸水'
    return ''


def get_person_info(simple_id = None):
    person_info = OrderedDict().fromkeys(["姓  名", "性  别", "年  龄", "病理诊断", "样本类型", "送检项目", "样本接收日期", "报告日期"], '')
    try:
        AF_API, MF_API, KF_API = set_api()
        AF = url_open(simple_id, AF_API)
        MF = url_open(simple_id, MF_API)
        KF = url_open(simple_id, KF_API)

        person_info["姓  名"] = AF["result"][0]["patientName"]
        person_info["性  别"] = AF["result"][0]["patientSexOptionText"]
        person_info["年  龄"] = str(AF["result"][0]["patientAge"])
        if str(AF["result"][0]["patientAge"]) == "0":
            person_info["年龄"] = ""
        person_info["病理诊断"] = AF["result"][0]["clinicalDiagnose"]
        person_info["样本类型"] = get_sample_type(AF["result"][0]["typeOptionValue"])
        # person_info["样本类型"] = AF["result"][0]["typeOptionText"]
        person_info["送检项目"] = 'SmartLung Plus'
        # person_info["报告版本"] = AF["result"][0]["reportVersionOptionValue"]
        # person_info["联系电话"] = KF["result"][0]["phone"]
        # person_info["送检编号"] = AF["result"][0]["number"]
        # person_info["送检单位"] = AF["result"][0]["customerName"]
        # person_info["样本采集日期"] = AF["result"][0]["collectTimeShow"]
        person_info["样本接收日期"] = MF["result"][0]["collectTimeShow"].split()[0]
        # person_info["用药史"] = AF["result"][0]["drug"]
        # person_info["癌种"] = AF["result"][0]["tumorTypeOptionText"]
        person_info["报告日期"] = time.strftime("%Y-%m-%d")
        # person_info["报告版本"] = AF["result"][0]["reportVersionOptionValue"]
    except Exception as e:
        print(e, file=sys.stderr)
    return person_info

def main():
    arg = sys.argv[1:]
    if not arg:
        print(__doc__)
        exit()
    dict_text = get_person_info(arg[0])
    print(dict_text)

if __name__ == "__main__":
    main()
