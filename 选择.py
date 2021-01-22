import pandas as pd
import numpy as np
def final():
    #处理wordij分词结果
    re = pd.read_excel('D:/hero time/swedn/wordij.xlsx')
    arr = pd.DataFrame(index=range(re.shape[0]),columns=['vertex1','vertex2','frequency'])
    for i in re.index:
        print(i)
        x,y=re.loc[i,'Pair']
        arr.loc[i,'vertex1'] = x
        arr.loc[i,'vertex2'] = y
        arr.loc[i,'frequency']=re.loc[i,'Frequency']
    arr.to_excel('D:/hero time/swedn/re.xlsx',index=False)

    #筛选与瑞典直接相关的词
    re = pd.read_excel('D:/hero time/swedn/re.xlsx')
    re = re[(re['vertex1']=="瑞典") | (re['vertex2']=="瑞典")]
    re.to_excel('D:/hero time/swedn/re_sw.xlsx', index=False)
    re = pd.read_excel('D:/hero time/swedn/re_sw.xlsx')
    re = re[re['vertex1'] != re['vertex2']]  # 去除自旋
    re_one = pd.DataFrame(index=range(re.shape[0]), columns=['vertex'])
    re_one['vertex'] = np.where(re['vertex1'] == "瑞典", re['vertex2'], re['vertex1'])
    re_two = pd.DataFrame(index=range(re_one['vertex'].unique().shape[0]), columns=['vertex'])  # 去冗余
    re_two['vertex'] = re_one['vertex'].unique()
    print(re_two)
    re_two.to_excel('D:/hero time/swedn/re_word.xlsx', index=False)

    #与瑞典相关且有共线关系的词对
    #未下载office需加载xlrd与apenpyxl模块
    re = pd.read_excel('D:/sci04/py/re_word.xlsx')
    vex = np.array(re['vertex'])
    origin = pd.read_excel('D:/sci04/py/re.xlsx')
    origin = origin[(origin['vertex1'] != '瑞典') & (origin['vertex2'] != '瑞典')]
    print(origin.shape)
    for i in origin.index:
        print(i)
        if((origin.loc[i,'vertex1'] in vex.tolist())==False or (origin.loc[i,'vertex2'] in vex.tolist())==False ):
           # print(origin.loc[i])
            origin = origin.drop(index = i)
    print(origin.shape)
    origin.to_csv('D:/sci04/py/re_pair.csv',encoding='utf-8',index=False)
    origin.to_excel('D:/sci04/py/re_pair.xlsx',index=False)
    #所有单个字的词
    one = pd.read_excel('D:/hero time/swedn/re_word.xlsx')
    for i in one.index:
        print(i)
        if (len(one.loc[i, 'vertex']) >= 2):
            one = one.drop(index=i)
    one.to_excel('D:/hero time/swedn/drop_one.xlsx', index=False)
    print(one)
    one = pd.read_excel('./swedn/drop_one.xlsx')
    one = one[one['len'] == 0]
    one.to_excel('./swedn/need.xlsx', index=False)
    last = pd.read_excel("./one/last.xlsx")
    last['label'] = np.where(last['label'] == 0, 0, 1)
    last.to_excel('./one/last.xlsx', index=False)
    last[last['label'] == 0].to_excel('./one/last_0.xlsx', index=False)
    last[last['label'] == 1].to_excel('./one/last_1.xlsx', index=False)
    print(last[last['label'] == 0].shape)
    #####################################
    root = pd.read_excel('./swedn/last_1.xlsx', engine='openpyxl')
    origin = pd.read_excel('./swedn/re.xlsx', engine='openpyxl')
    for i in origin.index:
        print(i)
        flag = 0
        for n in root.index:
            if (root.loc[n, 'vertex'] == origin.loc[i, 'vertex1']):
                flag += 1
        for m in root.index:
            if (root.loc[m, 'vertex'] == origin.loc[i, 'vertex2']):
                flag += 1
        if flag < 2:
            origin = origin.drop(index=i)
    origin.to_excel("./swedn/re_g.xlsx", index=False)
    print(origin.shape)
    ########################################################3
    root = pd.read_excel('./swedn/re_g.xlsx', engine='openpyxl')
    origin = pd.read_excel('./swedn/drop_re.xlsx', engine='openpyxl')
    print(root.shape)
    for i in root.index:
        print(i)
        for n in origin.index:
            if (root.loc[i, 'vertex1'] == origin.loc[n, 'word']):
                if origin.loc[n, 'replace'] == 1:
                    root = root.drop(index=i)
                    break
                else:
                    root.loc[i, 'vertex1'] = origin.loc[n, 'replace']
            if (root.loc[i, 'vertex2'] == origin.loc[n, 'word']):
                if origin.loc[n, 'replace'] == 1:
                    root = root.drop(index=i)
                    break
                else:
                    root.loc[i, 'vertex2'] = origin.loc[n, 'replace']
    print(root)
    print(root.shape)
    for i in root.index:
        print(i)
        if (root.loc[i, 'vertex2'] == root.loc[i, 'vertex1']):
            root = root.drop(index=i)

    root.to_excel('./swedn/re_g2.xlsx', index=False)
    ##################################################################
    group = pd.read_excel("./group/group.xlsx", engine='openpyxl')
    origin = pd.read_excel("./swedn/re_g2.xlsx", engine='openpyxl')
    decision = pd.read_excel("./group/decision.xlsx", engine='openpyxl')
    print(decision)
    for i in decision.index:
        if (decision.loc[i, 'decision'] == 1):
            # print(origin)
            origin_1 = origin
            group_1 = group[group['Group'] == 'G{}'.format(decision.loc[i, 'number'])]
            for k in origin_1.index:
                print(k)
                flag = 0
                for n in group_1.index:
                    if (group_1.loc[n, 'Vertex'] == origin_1.loc[k, 'vertex1']):
                        flag += 1
                        break
                for m in group_1.index:
                    if (group_1.loc[m, 'Vertex'] == origin_1.loc[k, 'vertex2']):
                        flag += 1
                        break
                if (flag < 2):
                    origin_1 = origin_1.drop(index=k)
            origin_1.to_excel("./group/G{}.xlsx".format(decision.loc[i, 'number']), index=False)
    ####################################################################
    group = pd.read_excel("./group/G1/G1_group.xlsx", engine='openpyxl')
    origin = pd.read_excel("./group/G1/G1_origin.xlsx", engine='openpyxl')
    for i in range(1, 12):
        origin_1 = origin
        # print(origin)
        group_1 = group[group['Group'] == 'G{}'.format(i)]
        for k in origin_1.index:
            print(k)
            flag = 0
            for n in group_1.index:
                if (group_1.loc[n, 'Vertex'] == origin_1.loc[k, 'vertex1']):
                    flag += 1
                    break
            for m in group_1.index:
                if (group_1.loc[m, 'Vertex'] == origin_1.loc[k, 'vertex2']):
                    flag += 1
                    break
            if (flag < 2):
                origin_1 = origin_1.drop(index=k)
        origin_1.to_excel("./group/G1/G{}.xlsx".format(i), index=False)
if __name__=="__main__":
    '''group = pd.read_excel("./group/G3/G3_group.xlsx", engine='openpyxl')
    origin = pd.read_excel("./group/G3/G3_origin.xlsx", engine='openpyxl')
    for i in range(1, 12):
        origin_1 = origin
        # print(origin)
        group_1 = group[group['Group'] == 'G{}'.format(i)]
        for k in origin_1.index:
            print(k)
            flag = 0
            for n in group_1.index:
                if (group_1.loc[n, 'Vertex'] == origin_1.loc[k, 'vertex1']):
                    flag += 1
                    break
            for m in group_1.index:
                if (group_1.loc[m, 'Vertex'] == origin_1.loc[k, 'vertex2']):
                    flag += 1
                    break
            if (flag < 2):
                origin_1 = origin_1.drop(index=k)
        origin_1.to_excel("./group/G3/G{}.xlsx".format(i), index=False)'''
    #for k in range(1,4):
    """for i in range(1,37):
        try:
            top = pd.read_excel("./group/G{}.xlsx".format(i),engine='openpyxl')
            print(i," OK")
        except:
            continue
        top.to_csv("./group_csv/G{}.csv".format(i),encoding='utf-8',index=False)"""


    for i in range(1,14):
        group = pd.read_excel("./second/G4/group.xlsx", engine='openpyxl')
        origin = pd.read_excel("./second/G4/origin.xlsx", engine='openpyxl')
        group = group[group['Group'] == 'G{}'.format(i)]
        print(origin)
        for k in origin.index:
            print(k)
            flag = 0
            for n in group.index:
                if (group.loc[n, 'Vertex'] == origin.loc[k, 'vertex1']):
                    flag += 1
                    break
            for m in group.index:
                if (group.loc[m, 'Vertex'] == origin.loc[k, 'vertex2']):
                    flag += 1
                    break
            if (flag < 2):
                origin = origin.drop(index=k)
        origin.to_excel("./second/G4/G{}.xlsx".format(i), index=False)
