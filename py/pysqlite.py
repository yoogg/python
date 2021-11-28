import sqlite3
#sqlite查询导入自动创建表脚本
class pysqlite:
    def __init__(self, dbPath='demo.db'):
        try:
            self.conn = sqlite3.connect(dbPath) # 链接数据库
            self.cursor = self.conn.cursor()
        except sqlite3.Error as e:
            print("数据库连接信息报错")
            raise e
    def dictFactory(self,cursor,row):
        """将sql查询结果整理成字典形式"""
        d={}
        for index,col in enumerate(cursor.description):
            d[col[0]]=row[index]
        return d
    def Query(self,sql:str)->list:
        """"""
        queryResult = self.conn.cursor().execute(sql).fetchall()
        return queryResult
    def selcet(self,sql:str)->dict:
        """调用该函数返回结果为字典形式"""
        self.conn.row_factory=self.dictFactory
        cur=self.conn.cursor()
        queryResult=cur.execute(sql).fetchall()
        return queryResult
    def insert(self,sql:str):
        print(f"执行的sql语句为\n{sql}")
        self.conn.cursor().execute(sql)
        self.conn.commit()
    def write(self, table_name, info_list):
        """
        根据table_name与info自动生成建表语句和insert插入语句
        :param table_name: 数据需要写入的表名
        :param info_list: 需要写入的内容，类型为列表
        :return:
        """
        sql_key = ''  # 数据库行字段
        sql_va=''
        sql_value = []  # 数据库值
        for value in info_list:
            sql_value.append(list(value.values())) 
        for key in info_list[0].keys():  # 生成insert插入语句
            sql_key = sql_key + ' ' + key + ','
            sql_va=sql_va+'?,'

        try:
            print(sql_key[:-1])
            self.cursor.executemany(
                "INSERT INTO %s (%s) VALUES (%s)" % (table_name, sql_key[:-1],sql_va[:-1]),sql_value)
            self.conn.commit()  # 提交当前事务
        except sqlite3.Error as e:
            print(str(e).split(':')[0])
            if str(e).split(':')[0] == "no such table":  # 当表不存在时，生成建表语句并建表
                sql_key_str = ''  # 用于数据库创建语句
                columnStyle = ''  # 数据库字段类型
                for key in info_list[0].keys():
                    sql_key_str = sql_key_str + ' ' + key + columnStyle + ','
                self.cursor.execute("CREATE TABLE %s (%s)" % (table_name, sql_key_str[:-1]))
                self.cursor.executemany(
                "INSERT INTO %s (%s) VALUES (%s)" % (table_name, sql_key[:-1],sql_va[:-1]),sql_value)
                self.conn.commit()  # 提交当前事务
            else:
                raise
    def Close(self):
        #关闭数据库连接
        self.conn.cursor().close()
        self.conn.close()
if __name__ == '__main__':
    import requests
    import json
    a=requests.get('https://www.layuiweb.com/test/table/demo3.json.js').text
    data=json.loads(a)['rows']['item']
    sql = DataTosqlite('ceshi.db')
    sql.write('ceshi', data)
    sql.Close

