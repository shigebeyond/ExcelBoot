
import pandas as pd

# Series
'''
tup=(1,2,3) # (元组)
s=pd.Series(tup)
print(s) # 不指定index, 则默认index为[0,1,len(s)-1]
print(s[1])
'''

# 删除行/列
# 初始化dataframe
df = pd.DataFrame({'a': ['a0', 'a1', 'a2'],
        'b': ['b0', 'b1', 'b2'],
        'c': ['c0', 'c1', 'c2']})
print("初始")
print(df)
print("\n")

print("第一行")
# 返回第一行
print(df.loc[0]) # 返回 Series
# 返回第二行
# print(df.loc[1])
print("\n")

# 删除第一行
print("删除第一行")
df_1 = df.drop(axis=0,index=0)
# df.drop(df.index[0], inplace=True)  # 删除第一行
# df.drop(df.index[0:3], inplace=True)  # 删除前三行
# df.drop(df.index[0, 2, 5], inplace=True)  # 删除第1行，第3行，第6行
print(df_1)
print("\n")

# 删除第一列和第二列
print("删除第一列和第二列")
df_2 = df.drop(axis=1,columns=['a','b'])
print(df_2)
print("\n")

# 删除第一列
print("删除第一列")
del df['a']
print(df)
print("\n")