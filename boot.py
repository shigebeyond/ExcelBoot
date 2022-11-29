#!/usr/bin/python3
# -*- coding: utf-8 -*-

import time
import sys
import os
import fnmatch
from pathlib import Path
from pyutilb.util import *
from pyutilb import log, ocr_youdao
import ast
from db import Db
import pandas as pd

# 跳出循环的异常
class BreakException(Exception):
    def __init__(self, condition):
        self.condition = condition # 跳转条件

# excel操作的基于yaml的启动器
class Boot(object):

    def __init__(self):
        # 步骤文件所在的目录
        self.step_dir = None
        # 已下载过的url对应的文件，key是url，value是文件
        self.downloaded_files = {}
        # 动作映射函数
        self.actions = {
            'sleep': self.sleep,
            'print': self.print,
            'for': self.do_for,
            'once': self.once,
            'break_if': self.break_if,
            'moveon_if': self.moveon_if,
            'moveon_if_exist_by': self.moveon_if_exist_by,
            'break_if_exist_by': self.break_if_exist_by,
            'break_if_not_exist_by': self.break_if_not_exist_by,
            'include': self.include,
            'set_vars': self.set_vars,
            'print_vars': self.print_vars,
            'connect_db': self.connect_db,
            'query_db': self.query_db,
            'export_excel': self.export_excel,
        }
        set_var('boot', self)
        # 当前文件
        self.step_file = None

    '''
    执行入口
    :param step_files 步骤配置文件或目录的列表
    '''
    def run(self, step_files):
        for path in step_files:
            # 1 模式文件
            if '*' in path:
                dir, pattern = path.rsplit(os.sep, 1)  # 从后面分割，分割为目录+模式
                if not os.path.exists(dir):
                    raise Exception(f'Step config directory not exist: {dir}')
                self.run_1dir(dir, pattern)
                return

            # 2 不存在
            if not os.path.exists(path):
                raise Exception(f'Step config file or directory not exist: {path}')

            # 3 目录: 遍历执行子文件
            if os.path.isdir(path):
                self.run_1dir(path)
                return

            # 4 纯文件
            self.run_1file(path)

    # 执行单个步骤目录: 遍历执行子文件
    # :param path 目录
    # :param pattern 文件名模式
    def run_1dir(self, dir, pattern ='*.yml'):
        # 遍历目录: https://blog.csdn.net/allway2/article/details/124176562
        files = os.listdir(dir)
        files.sort() # 按文件名排序
        for file in files:
            if fnmatch.fnmatch(file, pattern): # 匹配文件名模式
                file = os.path.join(dir, file)
                if os.path.isfile(file):
                    self.run_1file(file)

    # 执行单个步骤文件
    # :param step_file 步骤配置文件路径
    # :param include 是否inlude动作触发
    def run_1file(self, step_file, include = False):
        # 获得步骤文件的绝对路径
        if include: # 补上绝对路径
            if not os.path.isabs(step_file):
                step_file = self.step_dir + os.sep + step_file
        else: # 记录目录
            step_file = os.path.abspath(step_file)
            self.step_dir = os.path.dirname(step_file)

        log.debug(f"Load and run step file: {step_file}")
        # 获得步骤
        steps = read_yaml(step_file)
        self.step_file = step_file
        # 执行多个步骤
        self.run_steps(steps)

    # 执行多个步骤
    def run_steps(self, steps):
        # 逐个步骤调用多个动作
        for step in steps:
            for action, param in step.items():
                self.run_action(action, param)

    '''
    执行单个动作：就是调用动作名对应的函数
    :param action 动作名
    :param param 参数
    '''
    def run_action(self, action, param):
        if 'for(' in action:
            n = self.parse_for_n(action)
            self.do_for(param, n)
            return

        if action not in self.actions:
            raise Exception(f'Invalid action: [{action}]')

        # 调用动作对应的函数
        log.debug(f"handle action: {action}={param}")
        func = self.actions[action]
        func(param)

    # --------- 动作处理的函数 --------
    # 当前url
    # @property
    # def current_url(self):
    #     if self.driver == None:
    #         return None
    #     return self.driver.current_url

    # 解析动作名中的for(n)中的n
    def parse_for_n(self, action):
        n = action[4:-1]
        # 1 数字
        if n.isdigit():
            return int(n)

        # 2 变量名, 必须是list类型
        n = replace_var(n, False)

        # pd.Series == None 居然返回pd.Series, 无语, 转为list
        if isinstance(n, pd.Series):
            return list(n)
        if n == None or not (isinstance(n, list) or isinstance(n, int)):
            raise Exception(f'Variable in for({n}) parentheses must be int or list or pd.Series type')
        return n

    # for循环
    # :param steps 每个迭代中要执行的步骤
    # :param n 循环次数/循环的列表
    def do_for(self, steps, n = None):
        label = f"for({n})"
        # 循环次数
        if n == None:
            n = sys.maxsize # 最大int，等于无限循环次数
            label = f"for(∞)"
        # 循环的列表值
        items = None
        if isinstance(n, list):
            items = n
            n = len(items)
        log.debug(f"-- For loop start: {label} -- ")
        last_i = get_var('for_i', False) # 旧的索引
        last_v = get_var('for_v', False) # 旧的元素
        try:
            for i in range(n):
                # i+1表示迭代次数比较容易理解
                log.debug(f"{i+1}th iteration")
                set_var('for_i', i+1) # 更新索引
                if items == None:
                    v = None
                else:
                    v = items[i]
                set_var('for_v', v) # 更新元素
                self.run_steps(steps)
        except BreakException as e:  # 跳出循环
            log.debug(f"-- For loop break: {label}, break condition: {e.condition} -- ")
        else:
            log.debug(f"-- For loop finish: {label} -- ")
        finally:
            set_var('for_i', last_i) # 恢复索引
            set_var('for_v', last_v) # 恢复元素

    # 执行一次子步骤，相当于 for(1)
    def once(self, steps):
        self.do_for(steps, 1)

    # 检查并继续for循环
    def moveon_if(self, expr):
        # break_if(条件取反)
        self.break_if(f"not ({expr})")

    # 跳出for循环
    def break_if(self, expr):
        val = eval(expr, globals(), bvars)  # 丢失本地与全局变量, 如引用不了json模块
        if bool(val):
            raise BreakException(expr)

    # 检查并继续for循环
    def moveon_if_exist_by(self, config):
        self.break_if_not_exist_by(config)

    # 跳出for循环
    def break_if_not_exist_by(self, config):
        if not self.exist_by_any(config):
            raise BreakException(config)

    # 跳出for循环
    def break_if_exist_by(self, config):
        if self.exist_by_any(config):
            raise BreakException(config)

    # 加载并执行其他步骤文件
    def include(self, step_file):
        self.run_1file(step_file, True)

    # 设置变量
    def set_vars(self, vars):
        for k, v in vars.items():
            v = replace_var(v)  # 替换变量
            set_var(k, v)

    # 打印变量
    def print_vars(self, _):
        log.info(f"Variables: {bvars}")

    # 睡眠
    def sleep(self, seconds):
        seconds = replace_var(seconds)  # 替换变量
        time.sleep(int(seconds))

    # 打印
    def print(self, msg):
        msg = replace_var(msg)  # 替换变量
        log.debug(msg)

    # 连接db
    def connect_db(self, config):
        self.db = Db(config['ip'], config['port'], config['dbname'], config['user'], config['password'], config['echo_sql'])

    # 查询db
    def query_db(self, config):
        for var, sql in config.items():
            sql = replace_var(sql)
            df = self.db.query_dataFrame(sql)
            set_var(var, df)

    # 导出excel
    def export_excel(self, config):
        # 获得导出的变量
        var = config['var']
        val = get_var(var)
        # list转DataFrame
        if not isinstance(val, pd.DataFrame):
            if not isinstance(val, list):
                raise Exception(f"变量[{var}]值不是DataFrame或list: {val}")
            if len(val) == 0:
                print(f"列表变量[{var}]为空, 不用导出excel")
                return
            fields = val[0].keys()
            val = pd.DataFrame(val, columns=fields)

        # 行号
        if 'rownum' in config and config['rownum']:
            self.add_rownum_col(val)

        # 获得导出的文件
        file = replace_var(config['file'])

        # 导出
        # print(val)
        print(f'导出excel: {file}')
        if file.endswith('csv'):
            val.to_csv(file)
        else:
            sheet = replace_var(config['sheet'])
            writer = pd.ExcelWriter(file)
            val.to_excel(writer, sheet, index=False)

    # 添加行号列
    def add_rownum_col(df):
        df.insert(0, '序号', range(1, 1 + len(df)))

# cli入口
def main():
    # 基于yaml的执行器
    boot = Boot()
    # 读元数据：author/version/description
    dir = os.path.dirname(__file__)
    meta = read_init_file_meta(dir + os.sep + '__init__.py')
    # 步骤配置的yaml
    step_files = parse_cmd('ExcelBoot', meta['version'])
    if len(step_files) == 0:
        raise Exception("Miss step config file or directory")
    try:
        # 执行yaml配置的步骤
        boot.run(step_files)
    except Exception as ex:
        log.error(f"Exception occurs: current step file is {boot.step_file}", exc_info = ex)
        raise ex

if __name__ == '__main__':
    main()