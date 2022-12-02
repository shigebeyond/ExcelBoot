from sqlalchemy import Column, ForeignKey, String, Integer, FLOAT, DATE, create_engine, desc, MetaData, Table
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import pandas as pd
import sys
import datetime
from pyutilb import log

# db类
class Db:
    def __init__(self, ip, port, dbname, user, password, echo_sql = True):
        self.ip = ip
        self.port = port
        self.dbname = dbname
        self.user = user
        self.password = password
        self.echo_sql = echo_sql

        # 引擎
        self.engine = create_engine(f'mysql+mysqlconnector://{user}:{password}@{ip}:{port}/{dbname}?charset=utf8', echo=echo_sql)

        # 元数据
        self.metadata = MetaData(self.engine)

        # 创建DBSession类型:
        DBSession = sessionmaker(bind=self.engine, autocommit=True)
        self.session = DBSession()
        self.session.expire_on_commit = False # commit时不过期旧对象, 不然每次commit都会导致大量旧对象重新加载, 反正一个业务方法commit后, 另一个业务方法会重新查询对象, 以便实现方法之间解耦

        # 表
        self.tables = {}

        # 上一条sql
        self.last_sql = ''

    # 获得表
    def get_table(self, name):
        if name not in self.tables:
            self.tables[name] = Table(name, self.metadata, autoload=True)

        return self.tables[name]

    # 插入表
    def insert(self, table_name, values):
        # 连接数据表
        table = self.get_table(table_name)

        # 构建insert语句
        insert_sql = table.insert()
        # print(insert_sql)

        self.exec_sql(insert_sql, values)

    # 预览sql
    def preview_sql(self, sql, params):
        # 只处理sql是字符串的情况
        if isinstance(sql, str):
            for k, v in params.items():
                if isinstance(v, str):
                    v = f"'{v}'"
                elif isinstance(v, (datetime.datetime, datetime.date)):
                    v = v.strftime("%Y%m%d")
                    v = f"'{v}'"
                sql = sql.replace(':' + k, str(v))

            log.debug(sql)

    # 查询sql
    def query_sql(self, sql, params={}):
        self.preview_sql(sql, params)
        self.last_sql = sql
        cursor = self.session.execute(sql, params=params)
        return cursor.fetchall()

    # 查询sql,并转为 DataFrame
    def query_dataFrame(self, sql, params={}):
        r = self.query_sql(sql, params)
        if (len(r) == 0):
            return pd.DataFrame()

        fields = r[0]._fields
        pd.set_option('display.max_columns', len(fields))
        pd.set_option('display.max_rows', 100)
        return pd.DataFrame(r, columns=fields)

    # 更新sql
    def exec_sql(self, sql, params={}):
        self.preview_sql(sql, params)
        self.last_sql = sql
        cursor = self.session.execute(sql, params=params)
        self.session.commit()
        # print(cursor.lastrowid)