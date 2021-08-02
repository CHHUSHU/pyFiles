import pandas as pd

data = pd.read_excel('C:\\Users\\xiaoy\\Documents\\pyFiles\\导入分析数据.xlsx')
a = data.loc[data['95d']>9.84,['样本']]
b = data.loc[data['94E']>19.5,['样本']]
c = data.loc[data['132A']>12,['样本']]
d = data.loc[data['90F']>18,['样本']]
e = data.loc[data[8601]>6.28,['样本']]
f = data.loc[data[2801]>9.93,['样本']]
g = data.loc[data[4203]>6.97,['样本']]

a_neg = a.loc[:232,['样本']]
b_neg = b.loc[:232,['样本']]
c_neg = c.loc[:232,['样本']]
d_neg = d.loc[:232,['样本']]
e_neg = e.loc[:232,['样本']]
f_neg = f.loc[:232,['样本']]
g_neg = g.loc[:232,['样本']]
list_neg=[]
for i in a_neg['样本']:
    list_neg.append(i)
for i in b_neg['样本']:
    list_neg.append(i)
for i in c_neg['样本']:
    list_neg.append(i)
for i in d_neg['样本']:
    list_neg.append(i)
for i in e_neg['样本']:
    list_neg.append(i)
for i in f_neg['样本']:
    list_neg.append(i)
for i in g_neg['样本']:
    list_neg.append(i)
l_neg = set(list_neg)
print('negaive:')
print(l_neg)
total_neg = 0
for j in l_neg:
    total_neg += 1
print(total_neg)

a_PSP = a.loc[233:264,['样本']]
b_PSP = b.loc[233:264,['样本']]
c_PSP = c.loc[233:264,['样本']]
d_PSP = d.loc[233:264,['样本']]
e_PSP = e.loc[233:264,['样本']]
f_PSP = f.loc[233:264,['样本']]
g_PSP = g.loc[233:264,['样本']]
list_PSP=[]
for i in a_PSP['样本']:
    list_PSP.append(i)
for i in b_PSP['样本']:
    list_PSP.append(i)
for i in c_PSP['样本']:
    list_PSP.append(i)
for i in d_PSP['样本']:
    list_PSP.append(i)
for i in e_PSP['样本']:
    list_PSP.append(i)
for i in f_PSP['样本']:
    list_PSP.append(i)
for i in g_PSP['样本']:
    list_PSP.append(i)
l_PSP = set(list_PSP)
print('PSP:')
print(l_PSP)
total_PSP = 0
for j in l_PSP:
    total_PSP += 1
print(total_PSP)

a_DLB = a.loc[265:288,['样本']]
b_DLB = b.loc[265:288,['样本']]
c_DLB = c.loc[265:288,['样本']]
d_DLB = d.loc[265:288,['样本']]
e_DLB = e.loc[265:288,['样本']]
f_DLB = f.loc[265:288,['样本']]
g_DLB = g.loc[265:288,['样本']]
list_DLB = []
for i in a_DLB['样本']:
    list_DLB.append(i)
for i in b_DLB['样本']:
    list_DLB.append(i)
for i in c_DLB['样本']:
    list_DLB.append(i)
for i in d_DLB['样本']:
    list_DLB.append(i)
for i in e_DLB['样本']:
    list_DLB.append(i)
for i in f_DLB['样本']:
    list_DLB.append(i)
for i in g_DLB['样本']:
    list_DLB.append(i)
l_DLB = set(list_DLB)
print('DLB:')
print(l_DLB)
total_DLB = 0
for j in l_DLB:
    total_DLB += 1
print(total_DLB)

a_PD = a.loc[289:295,['样本']]
b_PD = b.loc[289:295,['样本']]
c_PD = c.loc[289:295,['样本']]
d_PD = d.loc[289:295,['样本']]
e_PD = e.loc[289:295,['样本']]
f_PD = f.loc[289:295,['样本']]
g_PD = g.loc[289:295,['样本']]
list_PD = []
for i in a_PD['样本']:
    list_PD.append(i)
for i in b_PD['样本']:
    list_PD.append(i)
for i in c_PD['样本']:
    list_PD.append(i)
for i in d_PD['样本']:
    list_PD.append(i)
for i in e_PD['样本']:
    list_PD.append(i)
for i in f_PD['样本']:
    list_PD.append(i)
for i in g_PD['样本']:
    list_PD.append(i)
l_PD = set(list_PD)
print('PD:' )
print(l_PD)
total_PD = 0
for j in l_PD:
    total_PD += 1
print(total_PD)

a_FTD = a.loc[296:304,['样本']]
b_FTD = b.loc[296:304,['样本']]
c_FTD = c.loc[296:304,['样本']]
d_FTD = d.loc[296:304,['样本']]
e_FTD = e.loc[296:304,['样本']]
f_FTD = f.loc[296:304,['样本']]
g_FTD = g.loc[296:304,['样本']]
list_FTD = []
for i in a_FTD['样本']:
    list_FTD.append(i)
for i in b_FTD['样本']:
    list_FTD.append(i)
for i in c_FTD['样本']:
    list_FTD.append(i)
for i in d_FTD['样本']:
    list_FTD.append(i)
for i in e_FTD['样本']:
    list_FTD.append(i)
for i in f_FTD['样本']:
    list_FTD.append(i)
for i in g_FTD['样本']:
    list_FTD.append(i)
l_FTD = set(list_FTD)
print('FTD:')
print(l_FTD)
total_FTD = 0
for j in l_FTD:
    total_FTD += 1
print(total_FTD)

a_AD = a.loc[305:,['样本']]
b_AD = b.loc[305:,['样本']]
c_AD = c.loc[305:,['样本']]
d_AD = d.loc[305:,['样本']]
e_AD = e.loc[305:,['样本']]
f_AD = f.loc[305:,['样本']]
g_AD = g.loc[305:,['样本']]
list_AD = []
for i in a_AD['样本']:
    list_AD.append(i)
for i in b_AD['样本']:
    list_AD.append(i)
for i in c_AD['样本']:
    list_AD.append(i)
for i in d_AD['样本']:
    list_AD.append(i)
for i in e_AD['样本']:
    list_AD.append(i)
for i in f_AD['样本']:
    list_AD.append(i)
for i in g_AD['样本']:
    list_AD.append(i)
l_AD = set(list_AD)
print('AD:')
print(l_AD)
total_AD = 0
for j in l_AD:
    total_AD += 1
print(total_AD)

neg = 232
特异性 = (neg-total_neg)/neg
PSP = 32
敏感度_PSP = total_PSP/PSP
DLB = 24
敏感度_DLB = total_DLB/DLB
PD = 7
敏感度_PD = total_PD/PD
FTD = 9
敏感度_FTD = total_FTD/FTD
AD = 128
敏感度_AD = total_AD/AD
准确度_AD = (total_AD + (neg-total_neg))/(AD+neg)
print('特异性:'+ str(特异性), '\n敏感度_PSP:'+ str(敏感度_PSP),'\n敏感度_DLB:'+ str(敏感度_DLB),'\n敏感度_PD:'+ str(敏感度_PD),'\n敏感度_FTD:'+ str(敏感度_FTD),'\n敏感度_AD:' + str(敏感度_AD))
print('准确度_AD')
print(准确度_AD)


