#!/usr/bin/env python3
# coding:utf-8

"""
usage:
        ddpcr_info 2017.02.23
        python3 ddpcr_info.py simple_id
        default:SERVER = "140.206.216.18"
                PORT = "8002"

        eg: python3 ddpcr_info.py S160000000
"""

import json
import sys
import io
import numpy as np
import urllib.request
from PIL import Image
from collections import OrderedDict

SERVER = "140.206.216.18"
PORT = "8002"

Trans_dict = {
    0:"阳性",
    1:"阴性"
}

def set_api():
    API = "http://{SERVER}:{PORT}/ddpcr/rest/commonDB/list/ddpcr/sampleNo=".format(SERVER=SERVER, PORT=PORT)
    return API

def get_ddpcr_info(simple_id = None):
    if simple_id == None:
        print("请输入正确的样本编号!")
        exit()
    temp_dict = OrderedDict()
    temp_list = []
    url = set_api()
    url = url + simple_id
    url_open = urllib.request.urlopen(url).read().decode("utf-8")
    json_info = json.loads(url_open)
    for sid in json_info['result'][0]['mutationVOMap'].keys():
        try:
            data = json_info['result'][0]['mutationVOMap'][sid]['fileByteArray']
            ret = io.BytesIO(bytes(np.array(data, dtype='uint8')))
            temp_dict["样本浓度(ng/ul)"] = json_info['result'][0]['concentration']
            temp_dict["上样体积(ul)"] = json_info['result'][0]['volume']
            temp_dict["突变型拷贝数(copise/20ul)"] = json_info['result'][0]['mutationVOMap'][sid]['mutantCopies']
            temp_dict["野生型拷贝数(copise/20ul)"] = json_info['result'][0]['mutationVOMap'][sid]['wildCopies']
            temp_dict["定性结果"] = Trans_dict[json_info['result'][0]['mutationVOMap'][sid]['result']]
            if sid == "E19Del":
                temp_dict["突变"] = "19del"
            else:
                temp_dict["突变"] = sid
            temp_dict["图片"] = Image.open(ret)
            temp_list.append(temp_dict)
            temp_dict = OrderedDict()
        except:
            continue
    return temp_list

def main():
    arg = sys.argv[1:]
    if not arg:
        print(__doc__)
        exit()
    dict_text = get_ddpcr_info(arg[0])
    print(dict_text)

if __name__ == "__main__":
    main()
