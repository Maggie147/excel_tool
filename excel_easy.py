# -*- coding:utf-8 -*-
from collections import OrderedDict
from collections import Iterable
from pyexcel_xls import save_data, get_data
import os, stat
import pprint


class ExcelHandle(object):
    def __init__(self):
        pass

    @staticmethod
    def full_path(fpath, fname):
        full_path = os.path.join(fpath, fname)
        if not os.path.exists(fpath):
            os.makedirs(fpath)
            os.chmod(fpath, stat.S_IRWXG)
        return full_path

    @staticmethod
    def get_default(data, tflag):
        if not data:
            tflag = tflag.lower()
            if tflag.startswith('h'):
                default = (u'序号', u'地址', u'资产名称', u'添加/发现时间', u'管理员', u'安全状态', u'可用性', u'备注')
            elif tflag.startswith('b'):
                default = ('id', 'address', 'name', 'found_time', 'administrator', 'safe_state', 'availability', 'description')
            else:
                default = tuple()
        else:
            default = tuple(data)
        return default

    @staticmethod
    def get_data_pyexcel(fname):
        excel = get_data(fname)
        for sheet_n in excel.keys():
            print "sheet name: ", sheet_n
            pprint.pprint(excel)
            print '\n'

    @staticmethod
    def save_data_pyexcel(fname, head, body):
        try:
            sheet_1 = []
            sheet_1.append(head)
            sheet_1.extend(body)
            sheet_data = OrderedDict()
            sheet_data.update({"Sheet1": sheet_1})
            save_data(fname, sheet_data) 
            return True         
        except Exception as e:
            print(e)
            return False

    @classmethod
    def save_xls_easy(cls, fpath, fname, data, head=None, data_keys=None):
        if not data and isinstance(data, Iterable):
            return False
        head = cls.get_default(head, 'h')
        data_keys = cls.get_default(data_keys, 'b')
        try:
            body = [[item.get(key).decode('utf-8') if isinstance(item.get(key), str) else item.get(key) for key in data_keys] for item in data]    
            if not body:
                return False
            full_path = cls.full_path(fpath, fname)
            return cls.save_data_pyexcel(full_path, head, body)
        except Exception as e:
            print(e)
            return False


def test():
    file_dir = './'
    fname = 'demo1.xls'
    head = [u'序号', u'地址', u'资产名称', u'添加/发现时间', u'管理员', u'安全状态', u'可用性', u'备注']
    data = [{'id': '1', 'address':'127.0.0.1', 'name':'bbb', 'found_time':'2012/03/06', 'administrator':'admin', 'safe_state':'safe1111', 'description':1}, 
            {'id': '2', 'address':'127.0.0.2', 'name':'cc', 'found_time':'2012/03/06', 'administrator':'safe1111', 'safe_state':'safe22', 'description':1, 'aa':'nihao'}]

    ret = ExcelHandle.save_xls_easy(file_dir, fname, data)
    if not ret:
        print("save xls file[{}] failed!".format(fname))
    else:
        print("save xls file[{}] success!".format(fname))


if __name__ == "__main__":
    test()
