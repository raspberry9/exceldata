#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
workbook은 엑셀로 작성된 데이터 시트를 로딩합니다.
로딩된 데이터는 named tuple 형태가 됩니다.
데이터를 찾기 위한 Key를 지정하면 해당 키로 더 빠르게 데이터를 찾을 수 있습니다.
시트명과 각 시트의 컬럼명에 python 변수명으로 사용할 수 없는 원소가 있으면 안됩니다.

* 엑셀 데이터 예시 *
================================
Key    StringValue NumericValue
--------------------------------
1       Hello	    1.50
2	    World	    1.50
================================
* 인덱스(1부터 시작)로 검사
> print data.FirstSheet[5].StringValue
>> 
* 인덱스 이외의 다른 값으로 검사
> print data.SecondSheet.findone(StringValue='Hello')
>>_(StringValue=u'Hello', NumericValue=1.5, Key=1.0)
> print data.SecondSheet.findmany(NumericValue=1.5)
>>[_(StringValue=u'Hello', NumericValue=1.5, Key=1.0), _(StringValue=u'World', NumericValue=1.5, Key=2.0)]

 - 이 코드는 python 2.7.11 버전 에서만 테스트 되었습니다.
 - 개선 요청 및 문의는 koo@kormail.net으로 해주세요.

Copyright (c) 2016, Jaseung Koo.
License: MIT (see LICENSE for details)
"""

from __future__ import with_statement

__author__ = 'Jaseung Koo'
__version__ = '0.0.1'
__license__ = 'MIT'

import re
import sys
import xlrd
from collections import namedtuple

def convert(v):
    if isinstance(v, dict):
        return namedtuple('_', v.keys())(**{x: convert(y) for x, y in v.items()})
    if isinstance(v, (list, tuple)):
        return [convert(x) for x in v]
    return v

SHEET_NAME_PATTERN = re.compile('^[\w_]+[A-Za-z0-9_]$')


class FilterList(list):
    def __init__(self, origin):
        super(FilterList, self).__init__(origin)

    def findone(self, **kwargs):
        for x in data.SecondSheet:
            if self._check(x, kwargs):
                return x
        return None

    def _check(self, data, filters):
        for key, value in filters.iteritems():
            if getattr(data, key) != value:
                return False
        return True
            
    def findmany(self, **kwargs):
        return filter(lambda x: self._check(x, kwargs), data.SecondSheet)


class Workbook(object):
    '''excel workbook'''
    def __init__(self, filename):
        workbook = xlrd.open_workbook(filename)
        for sheet_name in workbook.sheet_names():
            if not SHEET_NAME_PATTERN.match(sheet_name):
                raise Exception("'%s' is not a valid sheet name. Please use alphanumeric characters only." % (sheet,))
            sheet = workbook.sheet_by_name(sheet_name)
            fields = sheet.row_values(0)
            data = {0: dict(zip(fields, [None for _ in range(len(fields))]))}
            for row in xrange(1, sheet.nrows):
                row_vals = sheet.row_values(row)
                data[row] = dict(zip(fields, row_vals))

            self.__dict__[sheet_name] = FilterList([convert(val)  for row, val in data.iteritems()])

if __name__ == '__main__':
    data = Workbook('Sample.xlsx')
    
    # 인덱스(1부터 시작)로 검사 - 5번 데이터의 StringValue 값
    print data.FirstSheet[5].StringValue
    
    # 특정 키 값으로 원소 한 개 찾기 - StringValue 값이 Hello 인 첫 번째 행
    hello = data.SecondSheet.findone(StringValue='Hello')
    print hello.Key, hello.StringValue, hello.NumericValue

    # 특정 키 값으로 원소 여러 개 찾기 - NumericValue가 1.5인 모든 행
    result_list = data.SecondSheet.findmany(NumericValue=1.5)
    print result_list[0].Key, result_list[0].StringValue, result_list[0].NumericValue
    print result_list[1].Key, result_list[1].StringValue, result_list[1].NumericValue
