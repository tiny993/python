#!/usr/bin/env python
# -*- coding: utf-8 -*-

from database import Database, DatabaseMethod
from metaclass import h_ptl_register, h_ptl_dlms_parse, h_ptl_dlms_pfobis
from excel import Excel


reg = h_ptl_register(
    PTL_TYPE='04010A00',
    DEVICE_TYPE=1,
    SCALAR=0,
    UNIT=0,
    RW='ro',
    EX_TYPE=0
)

parse = h_ptl_dlms_parse(
    PTL_TYPE='04010A00',
    DEVICE_TYPE=1,
    SCALAR=0,
    UNIT=0,
    RW='ro',
    IS_DEFAULT=1
)

pfobis = h_ptl_dlms_pfobis(
    PTL_TYPE='04010A00',
    DEVICE_TYPE=1,
    RW='ro',
    IS_DEFAULT=1,
    IS_SELECT=1
)

#  1: oracle 2: mysql 3: postgre
#d = Database(DatabaseMethod.Oracle)
#d.connect('ami', 'ami', 'ami', '10.32.233.179', '1521')

models = [reg, pfobis, parse]

#   1. 配置需要修改项目类型，也就是Excel的第二个参数（0:cetus, 1:ruby）
#   2. 使用在excel.py 320行配置支持的设备类型

#   project type, 0:cetus, 1:ruby
e = Excel(r'D:\OBIS_Template.xlsx', 1)
for model in models:
    if isinstance(model, h_ptl_register):
        lst = e.parse(r'register', model)
        e.save(lst, 'D:\h_ptl_register.sql')
    elif isinstance(model, h_ptl_dlms_pfobis):
        lst = e.parse(r'pfobis', model)
        e.save(lst, 'D:\h_ptl_dlms_pfobis.sql')
    elif isinstance(model, h_ptl_dlms_parse):
        lst = e.parse(r'parse', model)
        e.save(lst, 'D:\h_ptl_dlms_parse.sql')

#d.insert(lst)

#d.close()
