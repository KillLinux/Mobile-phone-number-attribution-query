import requests,json
import pandas as pd
print('手机号归属地查询 Developer by Dr.KillLinux')
pn_lo=str(input('原始文件路径（绝对路径）：'))
pnl_index=str(input('抓取的列名：'))
n_index=int(input('插入新表格的索引位置：'))
n_name=str(input('插入新表格的列名：'))
n_sheet=str(input('新表格的sheet名：'))
n_lo=str(input('新文件路径（绝对路径）：'))
print('开始处理')
pn=pd.read_excel(pn_lo) 
pnl = pn[pnl_index].tolist()
c=0
for i in pnl:
    pnl[c]=str(i)
    c+=1
gsd=[]
print('抓取数据 API By 360')
for i in (pnl):
    a=json.loads(requests.get('https://cx.shouji.360.cn/phonearea.php?number='+i).text)
    a=a['data']
    gsd.append(a['province']+a['city'])
pn.insert(n_index,n_name,gsd)
pn.to_excel(n_lo,sheet_name=n_sheet,index=False)
print('新文件在'+n_lo)