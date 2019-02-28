#!/usr/bin/env python
# -*- coding: utf-8 -*-
""" Simple ORM using metaclass """


class Field(object):

    def __init__(self, name, column_type):
        self.name = name
        self.column_type = column_type

    def __str__(self):
        return '<%s:%s>' % (self.__class__.__name__, self.name)


class CharField(Field):
    def __init__(self, name, m_len):
        char_len = 'char(' + str(m_len) + ")"
        super(CharField, self).__init__(name, char_len)


class StringField(Field):
    def __init__(self, name, m_len):
        str_len = 'varchar(' + str(m_len) + ")"
        super(StringField, self).__init__(name, str_len)


class DateField(Field):
    def __init__(self, name):
        super(DateField, self).__init__(name, 'date')


class IntegerField(Field):
    def __init__(self, name):
        super(IntegerField, self).__init__(name, 'bigint')


class ModelMetaclass(type):
    def __new__(mcs, name, bases, attrs):
        if name == 'Model':
            return type.__new__(mcs, name, bases, attrs)
        mappings = dict()
        for k, v in attrs.items():
            if isinstance(v, Field):
                mappings[k] = v
        for k in mappings.keys():
            attrs.pop(k)
        attrs['__mappings__'] = mappings
        attrs['__table__'] = name
        return type.__new__(mcs, name, bases, attrs)


class Model(dict, metaclass=ModelMetaclass):
    def __init__(self, **kw):
        super(Model, self).__init__(**kw)

    def __getattr__(self, key):
        try:
            return self[key]
        except KeyError:
            raise AttributeError(r"'Model' object has no attribute '%s'" % key)

    def __setattr__(self, key, value):
        self[key] = value

    def sql(self):
        fields = []
        args = []
        for k, v in self.__mappings__.items():
            fields.append(v.name)
            args.append(getattr(self, k, None))

        s = val2str(args)
        sql = 'insert into %s (%s) values (%s)' % (self.__table__, ','.join(fields), s)
        return sql


def val2str(args):
    val_str = ''
    idx = 1
    for val in args:
        if val is None:
            val_str = val_str + 'null'
        elif isinstance(val, int):
            val_str = val_str + str(val)
        else:
            val_str = val_str + '\'' + str(val) + '\''

        if idx != len(args):
            val_str = val_str + ', '
        idx = idx + 1

    return val_str


class h_ptl_register(Model):
    PTL_TYPE = StringField('PTL_TYPE', 30)
    DEVICE_TYPE = IntegerField('DEVICE_TYPE')
    REGISTER_ID = StringField('REGISTER_ID', 100)
    PROTOCOL_ID = StringField('PROTOCOL_ID', 30)
    CLASS_ID = IntegerField('CLASS_ID')
    ATTRIBUTE_ID = IntegerField('ATTRIBUTE_ID')
    REGISTER_DESC = StringField('REGISTER_DESC', 100)
    SCALAR = IntegerField('SCALAR')
    UNIT = IntegerField('UNIT')
    UOM_NAME = StringField('UOM_NAME', 20)
    STD_UOM = StringField('STD_UOM', 20)
    RW = CharField('RW', 2)
    DATA_TYPE = IntegerField('DATA_TYPE')
    CLIENT_ID = IntegerField('CLIENT_ID')
    A = IntegerField('A')
    B = IntegerField('B')
    C = IntegerField('C')
    D = IntegerField('D')
    E = IntegerField('E')
    F = IntegerField('F')
    DATA_CLASS = StringField('DATA_CLASS', 20)
    FUNCTION_NAME = StringField('FUNCTION_NAME', 50)
    SHORT_NAME = StringField('SHORT_NAME', 10)
    EX_TYPE = IntegerField('EX_TYPE')


class h_ptl_dlms_parse(Model):
    PTL_TYPE = StringField('PTL_TYPE', 30)
    DEVICE_TYPE = IntegerField('DEVICE_TYPE')
    PROFILE_OBIS = StringField('PROFILE_OBIS', 30)
    CAPTURE_OBIS = StringField('CAPTURE_OBIS', 30)
    ATTRIBUTE_ID = IntegerField('ATTRIBUTE_ID')
    CLASS_ID = IntegerField('CLASS_ID')
    SCALAR = IntegerField('SCALAR')
    UNIT = IntegerField('UNIT')
    UOM_NAME = StringField('UOM_NAME', 20)
    STD_UOM = StringField('STD_UOM', 20)
    DATA_CLASS = StringField('DATA_CLASS', 20)
    RW = CharField('RW', 2)
    OBIS_IDX = IntegerField('OBIS_IDX')
    REMARK = StringField('REMARK', 100)
    IS_DEFAULT = IntegerField('IS_DEFAULT')
    EX_METER_TYPE = IntegerField('EX_METER_TYPE')
    SAVE_TSTMP = DateField('SAVE_TSTMP')


class h_ptl_dlms_pfobis(Model):
    PTL_TYPE = StringField('PTL_TYPE', 30)
    DEVICE_TYPE = IntegerField('DEVICE_TYPE')
    PROFILE_OBIS = StringField('PROFILE_OBIS', 30)
    PROFILE_NAME = StringField('PROFILE_NAME', 130)
    CLASS_ID = IntegerField('CLASS_ID')
    ATTRIBUTE_ID = IntegerField('ATTRIBUTE_ID')
    RW = CharField('RW', 2)
    PROFILE_TYPE = IntegerField('PROFILE_TYPE')
    SHORT_NAME = StringField('SHORT_NAME', 50)
    EX_METER_TYPE = IntegerField('EX_METER_TYPE')
    IS_DEFAULT = IntegerField('IS_DEFAULT')
    MTR_DATA_OBIS = StringField('MTR_DATA_OBIS', 30)
    DATA_INTERVAL = IntegerField('DATA_INTERVAL')
    IS_SELECT = IntegerField('IS_SELECT')
