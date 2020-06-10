# -*- coding:utf-8 -*-
import json

import requests


j =""
def json_to_dict():
    dict = json.loads(s=j)
    actList = dict['result']['result']
    return  actList# {'id': '007', 'name': '007', 'age': 28, 'sex': 'male', 'phone': '13000000000', 'email': '123@qq.com'}


if __name__ == '__main__':
    # actList = json_to_dict()
    # pos = 0
    # print("活动名称： "+actList[pos]["activity_name"])
    # print("商品明细")
    # print(actList[pos]["activity_sku_info"])
    target_url = "https://mms.pinduoduo.com/lakemms/activityGoods/list"
    cookie =""

    headers = {
    "accept": "application/json",
    "accept-encoding": "gzip, deflate, br",
    "accept-language": "zh-CN, zh;q = 0.9,es;q = 0.8,en;q = 0.7",
    "content-length": "204",
    "content-type": "application/json;charset:UTF-8",
    "cookie": cookie,
    "origin": "https://mms.pinduoduo.com",
    "referer": "https://mms.pinduoduo.com/act/register_record",
    "user-agent":"Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 75.0.3770.80 Safari / 537.36"
    }

    page_size = 10
    page_number = 1
    request_data = {"status":[101,102,103,104,105,106,107,201,202,203,301,302,401,402,501,502,601,602,701,702],"page_size":page_size,"page_number":page_number,"order_by":"created_at","sort_by":"desc","is_wait_handle_invite_cut_price":'false'}


    data_response = requests.post(target_url, data=request_data, headers=headers)
    print(data_response.text)
