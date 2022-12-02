#!/usr/bin/python3
# -*- coding: utf-8 -*-

from pyutilb.util import *
import pandas as pd

# DataFrame的列变换器
class DfColMapper(object):

    def __init__(self, df: pd.DataFrame):
        self.df = df

    # 变换列
    def map(self, col, expr):
        if '(' in expr:  # 1 函数调用, 如 link_sheet(目录)
            self.parse_and_call_col_func(col, expr)
            return

        # 2 表达式：直接eval
        self.call_eval(col, expr)

    # 执行列表达式: 解析与调用函数
    def parse_and_call_col_func(self, col, expr):
        # 1 解析函数与参数
        if '(' in expr:
            func, params = parse_func(expr)
        else:
            func = expr
            params = []

        # 2 调用函数
        # 2.1 add_id/rm函数单独处理
        if func == 'rm':
            self.rm(col)
            return
        if func == 'add_id':
            self.add_id(col)
            return

        # 2.2 调用系统函数
        # 逐行调用函数, 来拼接列的每个值
        r = []
        for row in self.df.itertuples():
            # 将[引用属性的参数]替换为属性值
            params2 = self.replace_attr_params(params, row)
            # 调用系统函数
            v = call_func(func, params2)
            r.append(v)
        self.df[col] = r

    # 读属性值，如果属性不存在，则读变量
    def get_attr_or_var(self, obj, attrname):
        # 1 有属性
        if hasattr(obj, attrname):
            return getattr(obj, attrname)  # 属性值

        # 2 无属性, 则读变量
        return get_var(attrname)

    # 将[引用属性的参数]替换为属性值
    def replace_attr_params(self, params, row):
        r = []

        # re正则匹配替换字符串 https://cloud.tencent.com/developer/article/1774589
        def replace(match, _ = None) -> str:
            attrname = match.group(1) # 属性名
            return self.get_attr_or_var(row, attrname)  # 属性值
        
        for i in range(0, len(params)):
            v = params[i]  # 参数
            if '$' in v:  # 参数是属性引用
                v = do_replace_var(v, replace=replace)
            r.append(v)
        return r

    # 执行列表达式: 直接eval
    def call_eval(self, col, expr):
        attrnames = re.findall(r'\$([\w\d_]+)', expr)  # 获得引用的属性名
        expr = expr.replace('$', '') # 干掉引用符

        # 逐行执行eval, 来拼接列的每个值
        r = []
        for row in self.df.itertuples():
            # 将[引用属性]作为eval的变量
            vars = self.build_attr_vars(attrnames, row)
            # eval
            v = eval(expr, vars)
            r.append(v)
        self.df[col] = r

    # 将[引用属性]作为eval的变量
    def build_attr_vars(self, attrnames, row):
        r = {}
        for attrname in attrnames:
            r[attrname] = self.get_attr_or_var(row, attrname)
        return r

    # 添加行号列: 需要单独调用
    def add_id(self, col):
        self.df.insert(0, col, range(1, 1 + len(self.df)))

    # 删除列: 需要单独调用
    def rm(self, col):
        del self.df[col]


