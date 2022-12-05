[GitHub](https://github.com/shigebeyond/ExcelBoot) | [Gitee](https://gitee.com/shigebeyond/ExcelBoot)

# ExcelBoot - yaml驱动Excel生成

## 概述
许多伙伴日常中有很多excel制作的重复性的工作，譬如
1. 将数据库中的表与字段导出为excel，作为数据结构交付文档；
2. 统计单张表的数据，并导出为excel
3. 统计几张表的数据，拼接起来，并导出为excel
4. 统计单个库的数据，并导出为excel 
5. 统计几个库的数据，拼接起来，并导出为excel 
6. 从几个excel中读取数据，加工，并导出为excel
7. 从json url中读取数据，加工，并导出为excel
8. 从库、excel、json url等多个数据源中读取数据，加工，并导出为excel
9. 各种样式调整，如列宽、行高、字体、颜色等等

这些excel制作的工作繁杂，而且重复性高，可以考虑通过写代码(python)方式来生成excel；

但是大部分伙伴开发能力不足，因此创作了ExcelBoot工具，支持通过yaml配置excel生成步骤（虽然不用写代码，但还是需要写yaml，还是有一定的学习门槛）；

框架通过编写简单的yaml, 就可以执行一系列复杂的excel操作步骤, 如查询数据/导出数据/列转换//提取变量/打印变量等，极大的简化了伙伴编写自动化测试脚本的工作量与工作难度，大幅提高人效；

框架通过提供类似python`for`/`if`/`break`语义的步骤动作，赋予伙伴极大的开发能力与灵活性，能适用于广泛的测试场景。

框架提供`include`机制，用来加载并执行其他的步骤yaml，一方面是功能解耦，方便分工，一方面是功能复用，提高效率与质量，从而推进测试整体的工程化。

## 特性
1. 底层excel操作基于 pandas 与 openpyxl 库来实现 
2. 使用 selenium-requests 扩展来处理post请求与上传请求
3. 支持通过yaml来配置执行的步骤，简化了生成代码的开发:
每个步骤可以有多个动作，但单个步骤中动作名不能相同（yaml语法要求）;
动作代表excel上的一种操作，如switch_sheet/export_df等等;
4. 支持类似python`for`/`if`/`break`语义的步骤动作，灵活适应各种场景
5. 支持`include`引用其他的yaml配置文件，以便解耦与复用

## todo
1. 支持更多的动作

## 安装
```
pip3 install ExcelBoot
```

安装后会生成命令`ExcelBoot`;

注： 对于深度deepin-linux系统，生成的命令放在目录`~/.local/bin`，建议将该目录添加到环境变量`PATH`中，如
```
export PATH="$PATH:/home/shi/.local/bin"
```

## 使用
```
# 1 执行单个文件
ExcelBoot 步骤配置文件.yml

# 2 执行多个文件
ExcelBoot 步骤配置文件1.yml 步骤配置文件2.yml ...

# 3 执行单个目录, 即执行该目录下所有的yml文件
ExcelBoot 步骤配置目录

# 4 执行单个目录下的指定模式的文件
ExcelBoot 步骤配置目录/step-*.yml
```

如执行 `ExcelBoot example/step-dbschema.yml`，输出如下
```
......
```
命令会自动操作并生成excel

## 步骤配置文件及demo
用于指定多个步骤, 示例见源码 [example](example) 目录下的文件;

顶级的元素是步骤;

每个步骤里有多个动作(如switch_sheet/export_df)，如果动作有重名，就另外新开一个步骤写动作，这是由yaml语法限制导致的，但不影响步骤执行。

简单贴出1个demo
1. 导出数据库中的表与字段: 详见 [example/step-dbschema.yml](example/step-dbschema.yml)
```yaml

```

## 配置详解
支持通过yaml来配置执行的步骤;

每个步骤可以有多个动作，但单个步骤中动作名不能相同（yaml语法要求）;

动作代表excel上的一种操作，如switch_sheet/export_df等等;

下面详细介绍每个动作:

1. print: 打印, 支持输出变量/函数; 
```yaml
# 调试打印
print: "总申请数=${dyn_data.total_apply}, 剩余份数=${dyn_data.quantity_remain}"
```

2. connect_db: 连接mysql数据库
```yaml
- connect_db:
    ip: 192.168.62.200
    port: 3306
    dbname: test
    user: root
    password: root
    echo_sql: true
```

3. start_edit: 开始编辑excel
```yaml
- start_edit: data/test数据结构.xlsx
```

4. end_edit: 结束编辑excel（保存）
```yaml
- end_edit:
```

5. switch_sheet: 切换sheet
```yaml
- switch_sheet: 目录
```

6. query_db: 查询sql, 并将查询结果放到变量中
```yaml
- query_db:
    # 查询结果放到变量tables
    tables: > # 查询sql
        SELECT
            TABLE_COMMENT as 表注释,
            TABLE_NAME as 表名
        FROM
            information_schema.TABLES
        WHERE
            TABLE_SCHEMA = 'test'
```

7. map_df_cols: DataFrame列转换，其中动作名中()包含的是list或DataFrame类型的变量名
```yaml
- map_df_cols(tables):
      # 新列名:每一行执行的函数表达式
      # 每一行的函数执行结果组成新列，表达式可以带变量，如 $列名，表示该行中指定列的值
      序号: add_id() # 加行号
      表链接: link_sheet($表名) # sheet链接
      搜索: link(链接, http://baidu.com?wd=$表名) # url链接
```

8. map_cols: 转换sheet列，相当于map_df_cols，差异在于map_df_cols转换的是变量，map_cols转换的是sheet
```yaml
- map_cols:
      header: true # 是否有表头，表示第一行作为列名
      # 新列名:每一行执行的函数表达式
      # 每一行的函数执行结果组成新列，表达式可以带变量，如 $列名，表示该行中指定列的值
      序号: add_id() # 加行号
      表链接: link_sheet($表名) # sheet链接
```

9. export_df: 将变量值导出到当前sheet
```yaml
- export_df: tables # 导出变量 tables 的值到当前sheet
```

10. export_db: 将sql查询结果导出到excel
```yaml
- export_db: select * from user # 查询sql
```

11. get_cell_value: 读取单元格的值
```yaml
- get_cell_value:
      # 变量msg取B2单元格的单个值
      msg: B2
      # 变量取list/tuple/set/pd.Series类型的值
      # 变量col_values取B列的多个值（一维数组）
      col_values: B
      # 变量row_values取第1行的多个值（一维数组）
      row_values: 1
      # 变量boud_values取B1到D2区域的多个值（二维数组）
      boud_values: B1:D2
```

12. set_cell_value: 设置单元格的值
```yaml
- set_cell_value:
      B2: txt
      # 值是变量表达式
      B4: $msg
      # 值是list/tuple/set/pd.Series类型的变量
      B: $col_values
      1: $row_values
```

13. cells: 遍历cell设置样式或值, 其中动作名中()包含的是范围字符串, 支持变量表达式
```yaml
- cells(A1:C2): # 指定区域的多个单元格
    # 设置每个单元格的样式
    fill: red
- cells(A): # 指定列的多个单元格
    fill: red
- cells(1): # 指定行的多个单元格
    fill: red
- cells(A1): # 单个单元格
    fill: red
```

cell独有的样式
```yaml
- cells(A1):
    value: test # 设置值
    style: Hyperlink  # 添加超连接样式
```

row/col/cell都支持的样式
```yaml
- cells(A1):
- cols(A):
- rows(1):
    # 填充颜色
    fill: red
    # 字体
    font:
        name: 宋体 # 字体类型
        size: 14 # 字体大小
        color: FFFF00 # 字体颜色
        bold: true # 是否加粗
        italic: true # 是否斜体
    # 对齐样式
    alignment:
        horizontal: center # 水平对齐模式
        vertical: center # 垂直对齐模式
        text_rotation: 45 # 旋转角度
        wrap_text: true # 是否自动换行
    # 边框
    border:
        style: thick # 边线样式：double/mediumDashDotDot/slantDashDot/dashDotDot/dotted/hair/mediumDashed/dashed/dashDot/thin/mediumDashDot/medium/thick
        color: FFFF0000 # 边线颜色
```

14. cols: 遍历col设置样式, 其中动作名中()包含的是范围字符串, 支持变量表达式
```yaml
- cols(D:E): # 多列
    fill: blue
- cols(D): # 单列
    fill: blue
```

col独有的样式
```yaml
- cols(D): # 单列
    width: 40
```

15. rows: 遍历row设置样式, 其中动作名中()包含的是范围字符串, 支持变量表达式
```yaml
- rows(4:5): # 多行
    fill: green
- rows(4): # 单行
    fill: green
```

row独有的样式
```yaml
- rows(4):
    height: 40
```

16. insert_rows: 插入行
```yaml
# 在第1行之上插入3行
- insert_rows: 1, 3
```

17. insert_cols: 插入列
```yaml
# 在第1列之前插入3列
- insert_cols: 1, 3
```

18. delete_rows: 删除行
```yaml
# 删除第1-4行
- delete_rows: 1, 3
```

19. delete_cols: 删除列
```yaml
# 删除第1-4列
- delete_cols: 1, 3
```

20. merge_cells: 合并单元格
```yaml
# 合并 C1 到 D2 区域的单元格
- merge_cells: C1:D2
```

21. unmerge_cells: 取消合并单元格
```yaml
# 取消合并 C1 到 D2 区域的单元格
- unmerge_cells: C1:D2
```

22. insert_image: 插入图片
```yaml
- insert_image: 
    # 在A1单元格处, 插入图片
    A1: a.png
    C1:
      image: c.txt
      size: 100,90 # 宽度,高度
```

23. insert_file: 插入文件
```yaml
- insert_file: 
    # 在A1单元格处, 插入文件a.txt
    A1: a.txt
    C1: c.txt
```

24. for: 循环; 
for动作下包含一系列子步骤，表示循环执行这系列子步骤；变量`for_i`记录是第几次迭代（从1开始）,变量`for_v`记录是每次迭代的元素值（仅当是list类型的变量迭代时有效）
```yaml
# 循环3次
for(3) :
  # 每次迭代要执行的子步骤
  - switch_sheet: test

# 循环list类型的变量urls
for(urls) :
  # 每次迭代要执行的子步骤
  - switch_sheet: test

# 无限循环，直到遇到跳出动作
# 有变量for_i记录是第几次迭代（从1开始）
for:
  # 每次迭代要执行的子步骤
  - break_if: for_i>2 # 满足条件则跳出循环
    switch_sheet: test
```

25. once: 只执行一次，等价于 `for(1)`; 
once 结合 moveon_if，可以模拟 python 的 `if` 语法效果
```yaml
once:
  # 每次迭代要执行的子步骤
  - moveon_if: for_i<=2 # 满足条件则往下走，否则跳出循环
    switch_sheet: test
```

26. break_if: 满足条件则跳出循环; 
只能定义在for/once循环的子步骤中
```yaml
break_if: for_i>2 # 条件表达式，python语法
```

27. moveon_if: 满足条件则往下走，否则跳出循环; 
只能定义在for/once循环的子步骤中
```yaml
moveon_if: for_i<=2 # 条件表达式，python语法
```

28. include: 包含其他步骤文件，如记录公共的步骤，或记录配置数据(如用户名密码); 
```yaml
include: part-common.yml
```

29. set_vars: 设置变量; 
```yaml
set_vars:
  name: shi
  password: 123456
  birthday: 5-27
```

30. print_vars: 打印所有变量; 
```yaml
print_vars:
```