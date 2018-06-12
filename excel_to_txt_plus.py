# -*- coding: utf-8 -*-
"""
Created on Wed May  9 18:14:45 2018

@author: lenovo
"""

# -*- coding: utf-8 -*- 
import xlrd,os

file_dir=r'E:\test'
target_dir=r'E:\test\target'

def open_excel(filename):
    try:
        data=xlrd.open_workbook(filename)
        return data
    except Exception as e:
        print (e)
def get_flile_list(file_dir):
    return os.listdir(file_dir)



def excel2txt(file_dir,byindex,split_str,suffix='.xls'):
   filelist=get_flile_list(file_dir)
   for file in filelist:
        basefile=os.path.splitext(file)
        target_file=basefile[0]+'.txt'
        if os.path.exists(target_dir+r'\\'+target_file):
            os.remove(target_dir+r'\\'+target_file)
        if basefile[1]==suffix:
            excel_data=open_excel(file_dir+'\\'+file)
            table=excel_data.sheets()[byindex]
            nrows=table.nrows
            ncols=table.ncols
            txtfile=open(target_dir+'\\'+target_file,'a+',encoding = 'utf-8')
            for i in range(1,nrows):
                for j in range(ncols):
                    if j==0:
                        cell_value=str(table.cell(i,j).value)
                    else:
                        cell_value=split_str+str(table.cell(i,j).value)
                    txtfile.write(cell_value)
                txtfile.write('\n')
            txtfile.close()
        
if __name__=='__main__':
    excel2txt(file_dir,0,'\t')