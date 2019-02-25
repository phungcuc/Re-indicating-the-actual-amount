import pandas as pd
import xlrd
import re

d = pd.ExcelFile('\Find No of Ctns.xlsx')
df = d.parse('sheet1', skiprows = 0)
for a in range(df.shape[0]):
    if not isinstance(df['Description'][a], basestring):
        df=df.drop([a])
df = df.reset_index()
del df['index']

Q = []
for a in df.Qty:
    aa = re.search('\d+',a)
    if aa:
        Q.append(aa.group(0))
        
des = []
for b in df.Qty:
    bb = re.search('\D+', b)
    if bb:
        des.append(bb.group(0))
        
u = []
for c in df.Description:
    cc = re.search('(\d+)(BOTTLE)', c)
    ccc = re.search('(\d+)[X]',c)
    if cc:
        u.append(cc.group(1))
    elif ccc:
        u.append(ccc.group(1))
    else:
        u.append('')

no_ctn = []
for i in range(df.shape[0]):
    if des[i] == 'CTN':
        no_ctn.append(Q[i])
    elif u[i] == '':
        f = int(Q[i]) / 6
        no_ctn.append(f)
    else:
        g = int(Q[i]) / int(u[i])
        no_ctn.append(g)
        
df['no_ctn'] = no_ctn
df.to_csv('result.csv')
