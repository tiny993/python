#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Excel operate"""
import re
from typing import List, Any

import xlrd
from metaclass import h_ptl_register, h_ptl_dlms_parse, h_ptl_dlms_pfobis

unit_map = {0: (None, None),
            27: ('W', 'kW'),
            28: ('VA', 'kVA'),
            29: ('var', 'kvar'),
            30: ('Wh', 'kWh'),
            31: ('VAh', 'kVAh'),
            32: ('varh', 'kvarh'),
            33: ('A', 'A'),
            34: ('C', 'C'),
            35: ('V', 'V'),
            36: ('Ua', 'Ua'),
            37: ('%', '%'),
            38: ('s', 's')}


unit_tran = {'KWH': 30,
             'KW': 27,
             'KVARH': 32,
             'KVAR': 29,
             'KVAH': 31,
             'KVA': 28,
             'VARH': 32,
             'A': 33,
             'C': 34,
             'V': 35,
             'UA': 36,
             '%': 37,
             'S': 38}


type_map = {'null-data': 0,
            'array': 1,
            'structure': 2,
            'boolean': 3,
            'bit-string': 4,
            'double-long': 5,
            'double-long-unsigned': 6,
            'octet-string': 9,
            'visible-string': 10,
            'utf8-string': 12,
            'bcd': 13,
            'integer': 15,
            'long': 16,
            'unsigned': 17,
            'long-unsigned': 18,
            'compact-array': 19,
            'long64': 20,
            'long64-unsigned': 21,
            'enum': 22,
            'float32': 23,
            'float64': 24,
            'date-time': 25,
            'date': 26,
            'time': 27,
            'dont-care': 255}


profile_map = {1: 'Daily',
               2: 'Monthly',
               3: 'AllLoadProfile',
               4: 'Event'}


def data_type(str_type):
    if str_type is '':
        return None

    s = re.compile(r'[\s\_]')
    str_type = s.sub('-', str_type)
    for x in type_map:
        if str_type.find(x) is not -1:
            return type_map[x]
    else:
        return None


def unit_type(str_type):
    vec = []
    if str_type is '':
        return 0

    if str_type.find('(') is not -1:
        vec = str_type.split('(')
        if len(vec) > 2:
            str_type = vec[2]
        else:
            str_type = vec[1]
    else:
        return 0

    for x in unit_tran:
        if str_type.upper().find(x) is not -1:
            return unit_tran[x]
    else:
        return 0


def pro_type(pro_name):
    if pro_name is '':
        return None

    if re.match(r'(D|d)onthly', pro_name):
        return profile_map[1]
    elif re.match(r'(M|m)onthly', pro_name):
        return profile_map[2]
    elif re.match(r'(P|p)rofile', pro_name):
        return profile_map[3]
    elif re.match(r'(E|e)vent', pro_name):
        return profile_map[4]
    else:
        return None


class Excel(object):
    _obj_head = ''
    _obj_class = 0
    _attr_idx = 0
    _OBIS = []
    _attr_flag = 0
    _proj_type = 0

    #   project type, 0:cetus, 1:ruby
    def __init__(self, path, projectType = 0):
        self._proj_type = projectType
        try:
            self._file = xlrd.open_workbook(path)

        except xlrd.error_text_from_code as e:
            print('Open excel file failed. Error %s' % e)

    @staticmethod
    def save(sqls, path):
        with open(path, 'w', encoding='utf-8', errors='ignore') as f:
            for x in sqls:
                f.write(x)
                f.write(';\n')
            f.write('commit;')


    def cetus_register(self, register_model, row):
        register_model.REGISTER_ID = self._obj_head + ' ' + row[5].replace('_', ' ')
        register_model.PROTOCOL_ID = '.'.join(self._OBIS)
        register_model.CLASS_ID = self._obj_class
        register_model.ATTRIBUTE_ID = self._attr_idx
        register_model.REGISTER_DESC = row[5].replace('_', ' ')

        if register_model.REGISTER_DESC.find('unit') is not -1:
            tmp = re.split(r'[\{\}\,\s]+', row[9])
            if len(tmp) > 3:
                register_model.SCALAR = tmp[1]
                register_model.UNIT = int(tmp[2])
                if unit_map.get(register_model.UNIT) is not None:
                    register_model.UOM_NAME = unit_map.get(register_model.UNIT)[0]
                    register_model.STD_UOM = unit_map.get(register_model.UNIT)[1]

        if data_type(row[6]) is not None:
            register_model.DATA_TYPE = data_type(row[6])
        else:
            register_model.DATA_TYPE = None

        if len(self._OBIS) is 6:
            register_model.A = int(self._OBIS[0])
            register_model.B = int(self._OBIS[1])
            register_model.C = int(self._OBIS[2])
            register_model.D = int(self._OBIS[3])
            register_model.E = int(self._OBIS[4])
            register_model.F = int(self._OBIS[5])

        return register_model.sql()

    def cetus_pfobis(self, pf_model):
        pf_model.PROFILE_OBIS = '.'.join(self._OBIS)
        pf_model.PROFILE_NAME = self._obj_head
        pf_model.CLASS_ID = self._obj_class
        pf_model.ATTRIBUTE_ID = self._attr_idx


    #
    #   ruby register
    #
    def ruby_register(self, register_model, row):
        register_model.REGISTER_ID = row[0].strip()
        register_model.PROTOCOL_ID = row[1].strip()
        register_model.CLASS_ID = int(row[2])
        register_model.ATTRIBUTE_ID = int(row[3])
        register_model.REGISTER_DESC = row[4].strip()
        register_model.UNIT = unit_type(row[5])
        if unit_map.get(register_model.UNIT) is not None:
            register_model.UOM_NAME = unit_map.get(register_model.UNIT)[0]
            register_model.STD_UOM = unit_map.get(register_model.UNIT)[1]

        obis = row[1].strip().split('.')
        if len(obis) is not None:
            register_model.A = obis[0]
            register_model.B = obis[1]
            register_model.C = obis[2]
            register_model.D = obis[3]
            register_model.E = obis[4]
            register_model.F = obis[5]

        if row[4].find('Capture Time') is not -1:
            register_model.DATA_CLASS = 'Capture time'

        if row[5] is not '':
            register_model.SHORT_NAME = row[5].split('(')[0].strip()

        return register_model.sql()


    #
    #   ruby_parse
    #
    def ruby_parse(self, parse_model, row, obis_idx, ex_meter_type):
        parse_model.DATA_CLASS = None
        parse_model.PROFILE_OBIS = row[0].strip()
        parse_model.CAPTURE_OBIS = row[1].strip()
        parse_model.ATTRIBUTE_ID = int(row[2])
        parse_model.CLASS_ID = int(row[3])
        parse_model.UNIT = unit_type(row[5])
        if unit_map.get(parse_model.UNIT) is not None:
            parse_model.UOM_NAME = unit_map.get(parse_model.UNIT)[0]
            parse_model.STD_UOM = unit_map.get(parse_model.UNIT)[1]

        if row[4] is not '':
            parse_model.REMARK = row[4].strip()
            if row[4].find('Capture Time') is not -1:
                parse_model.DATA_CLASS = 'Capture time'

            elif row[4].find('Clock') is not -1:
                parse_model.DATA_CLASS = 'Date Time'
                self.obis_idx = -1

        parse_model.OBIS_IDX = obis_idx
        parse_model.EX_METER_TYPE = ex_meter_type

        return parse_model.sql()


    def ruby_pfobis(self, pfobis_model, row, ex_meter_type):
        pfobis_model.PROFILE_OBIS = row[0].strip()
        pfobis_model.PROFILE_NAME = row[1].strip()
        pfobis_model.CLASS_ID = int(row[2])
        pfobis_model.ATTRIBUTE_ID = int(row[3])
        pfobis_model.PROFILE_TYPE = int(row[4])
        pfobis_model.SHORT_NAME = row[5].strip()
        pfobis_model.EX_METER_TYPE = ex_meter_type
        if row[6] is not '':
            pfobis_model.DATA_INTERVAL = int(row[6])
        else:
            pfobis_model.DATA_INTERVAL = None
        return pfobis_model.sql()


    #
    #   parse cetus
    #
    def cetus(self, sheet, rows, model):
        str_lst = []
        for x in range(rows):
            row = sheet.row_values(x)
            if self._attr_flag is 0:
                if x < 3:  # ignore the first three lines
                    continue

                if isinstance(row[7], float):
                    row[7] = int(row[7])

                if row[7] is not '' and str(row[7]).isdigit():  # object
                    self._obj_class = int(row[7])

                    if row[5] is not '' and row[5].find('\n'):  # only get first line
                        self._obj_head = row[5].split('\n')[0]
                    else:
                        self._obj_head = row[5].replace('_', ' ')
                else:
                    continue

                if row[9] is not '':
                    self._OBIS = re.split(r'[\s\-\:\.]+', row[9])
                else:
                    continue

                self._attr_flag = 1
                continue
            else:
                if row[4] is not '' and self._attr_idx < int(row[4]):
                    self._attr_idx = int(row[4])
                    str_lst.append(self.cetus_register(model, row))
                else:
                    self._obj_head = ''
                    self._obj_class = 0
                    self._attr_idx = 0
                    self._OBIS = []
                    self._attr_flag = 0

        return str_lst


    #
    #   parse ruby
    #
    def ruby(self, name, sheet, rows, model):
        str_lst = []
        idx = 0
        for x in range(rows):
            if x is 0:
                continue
            row = sheet.row_values(x)

            if row[1] is '':
                continue

            if name is 'register':
                str_lst.append(self.ruby_register(model, row))

            #   10: single phase meter
            #   11: three-pahse four-wire meter
            #   14: CT meter
            meter_type = [10, 11, 14]
            if name is 'pfobis':
                for i in meter_type:
                    str_lst.append(self.ruby_pfobis(model, row, i))

            if name is 'parse':
                if row[4].find('Clock') is not -1:
                    idx = 0

                #   sp
                if row[6] is not '':
                    idx = idx + 1
                    str_lst.append(self.ruby_parse(model, row, idx, 10))

                #   pp
                if row[7] is not '':
                    idx = idx + 1
                    str_lst.append(self.ruby_parse(model, row, idx, 11))

        return str_lst


    def parse(self, sheet_name, model):
        sheet = self._file.sheet_by_index(self._file.sheet_names().index(sheet_name))

        rows = sheet.nrows

        if self._proj_type == 0:
            return self.cetus(sheet, rows, model)
        else:
            return self.ruby(sheet_name, sheet, rows, model)
