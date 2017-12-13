# -*- coding: utf-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
import csv
import pandas as pd

class Generate_csv(object):
    def __init__(self, count, directory, by):
        self.count = count
        self.directory = directory
        self.by = by

    def data_lst(self):
        for idx in range(0,self.count):
            if idx == 0:
                file_name = self.directory +self.by+"%28진료년월%29.xls"
            else:
                file_name = self.directory +self.by+'%28진료년월%29 ('+str(idx)+').xls'

            workbook = xlrd.open_workbook(file_name)
            worksheet_index = workbook.sheet_by_index(0)
            num_rows = worksheet_index.nrows
            num_cols = worksheet_index.ncols

            row_val = []
            for row_num in range(num_rows):
                row_val.append(worksheet_index.row_values(row_num))

            if idx == 0 :
                rows = row_val[6:]
            else :
                rows = row_val[8:]

            try:
                self.write_csv(rows, idx)
                self.split_csv()
            except Exception as e:
                print (e)
                continue

    def write_csv(self, rows, idx):
        if idx == 0 :
            with open("csv_file/mdfee_data"+self.by+".csv","wt") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
        else:
            with open("csv_file/mdfee_data"+self.by+".csv","a") as f:
                writer = csv.writer(f)
                writer.writerows(rows)

    def split_csv(self) :
        df_crawl = pd.read_csv("csv_file/mdfee_data"+self.by+".csv")
        df_crawl = df_crawl.drop(0,0)
        df_crawl.rename(columns={"Unnamed: 0":"mdfeeCd","진료년월" : "locationNm"}, inplace = True)
        locaNm_lst = df_crawl['locationNm'].unique().tolist()
        code_lst = df_crawl['mdfeeCd'].unique().tolist()
        df_new = pd.DataFrame(index=[[j for j in code_lst for i in range(len(locaNm_lst))],[i for i in locaNm_lst] * len(code_lst)])
        df_new = df_new.reset_index(drop=False )

        # 컬럼명 변경
        df_new.rename(columns={"level_0":"mdfeeCd","level_1":"locationNm"}, inplace = True)

        df_merge = pd.merge(df_crawl , df_new , how = 'outer' , on=['mdfeeCd','locationNm'] )
        #df_merge = df_merge.sort_values(by = ['mdfeeCd','locationNm'])
        df_merge_lst = list(df_merge.columns)

        # csv 파일 분할
        for num in range(1,4):
            col_lst = [df_merge_lst[2:][i] for i in range(num,len(df_merge_lst[2:]),3)]
            col_lst.insert(0, 'mdfeeCd')
            col_lst.insert(1, 'locationNm')
            if num == 1:
                df_merge[col_lst].to_csv('csv_file/Number_of_patients'+self.by+'.csv',encoding='utf-8')
            elif num == 2 :
                df_merge[col_lst].to_csv('csv_file/Total_usage'+self.by+'.csv',encoding='utf-8')
            elif num == 3 :
                df_merge[col_lst].to_csv('csv_file/Amount_of_treatment'+self.by+'.csv' ,encoding='utf-8')


        
