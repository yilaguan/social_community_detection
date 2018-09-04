# coding=utf-8

# Created by JetBrains Pycharm
# @Time      : 2018/9/4 13:59
# @Author    : zhangliang
# @File      : syndrome.py
# @Email     : zhangliangxgd@163.com

class Syndrome(object):
    def __init__(self, name, column_number, medicine_start, medicine_end, threshold, file_name):
        self.name = name
        self.column_number = column_number
        self.medicine_start = medicine_start
        self.medicine_end = medicine_end
        self.threshold = threshold
        self.file_name = file_name

    def get_name(self):
        return self.name

    def get_column_number(self):
        return self.column_number

    def get_medicine_start(self):
        return self.medicine_start

    def get_medicine_end(self):
        return self.medicine_end

    def get_threshold(self):
        return self.threshold

    def get_file_name(self):
        return self.file_name
