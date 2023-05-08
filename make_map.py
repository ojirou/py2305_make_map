import os
import datetime as dt
import subprocess
import shutil
import pandas as pd
datestr='230507'
# file_org='data_db//master_db'+datestr+'.xlsx'
# file_exl='master_list'+datestr+'.xlsx'
file_db='list'+datestr+'.xlsx'
file_map='map'+datestr+'.xlsx'
file_map2='map_r'+datestr+'.xlsx'
df_list=pd.read_excel(file_db, sheet_name='Sheet1', header=0)
# df_master.to_excel(file_exl)
years=['2020', '2021', '2022', '2023']
df_map=pd.DataFrame({
    "組織":["ISO", "IEC"],
})
for year in years:
    year=int(year)
    year_minus1=int(year)-1
    year_plus1=int(year)+1
    year_str=str(year)+"年"
    file_db_1year=year_str+file_db
    df_list_1year=df_list[(df_list['年月'] > dt.datetime(year-1,12,31)) & (df_list['年月'] < dt.datetime(year_plus1,1,1))]
    df_list_1year=df_list_1year[['組織', '記号']]
    df_map=df_map.merge(df_list_1year, on='組織', how='left')
    df_list_1year[:0]
df_map.columns=['組織', '2020年', '2021年', '2022年', '2023年']
df_map.to_excel(file_map)
# subprocess.Popen(["start", "", file_map], shell=True)
df_map2=pd.DataFrame(columns=['組織', '2020年', '2021年', '2022年', '2023年'])
orgs=["ISO", "IEC"]
for org in orgs:
    df_org=df_map[df_map['組織'].isin([org])]
    unique_list=[df_org[column_name].unique().tolist() for column_name in ['2020年', '2021年', '2022年', '2023年']]
    df_org=pd.DataFrame(unique_list).T
    df_org.columns=['2020年', '2021年', '2022年', '2023年']
    df_org['組織']=org
    df_org=df_org.reindex(columns=['組織', '2020年', '2021年', '2022年', '2023年'])
    df_map2=df_map2.append(df_org)
df_map2.to_excel(file_map2)
subprocess.Popen(["start", "", file_map2], shell=True)