
import pandas as pd

tup=(1,2,3) # (元组)
s=pd.Series(tup)
print(s) # 不指定index, 则默认index为[0,1,len(s)-1]
print(s[1])