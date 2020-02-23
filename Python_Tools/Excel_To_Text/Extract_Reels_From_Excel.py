# -*- coding: utf-8 -*-
'''
Author: Mingjun Wu
Date: 02/20/2020
'''
from __future__ import print_function
import pandas as pd
import os
import json
import sys
def ExcelToTxt(params):
    '''
    Convert Excel into Text
    '''
    print('Converting... ...')
    ExcelName = params['ExcelName']
    SheetNames = params['SheetNames']
    SheetRange = params['SheetRange']
    OutTxt = params['OutTxtName']
    NumSheet = len(SheetNames)
    if(len(SheetRange)!=NumSheet):
        print('The number of SheetNames is not equal to the number of SheetRange')
        return
    print("Excel Name:", ExcelName)
    print("There are {} sheets".format(NumSheet))
    with open(OutTxt,'w') as fout:
        for i, Sheet in enumerate(SheetNames):
            print("Sheet: {}, Range: {}".format(Sheet, SheetRange[i]))
            SRange = SheetRange[i][2]+':'+SheetRange[i][3]
            print(SRange)
            print('Reading... ...')
            current_sheet = pd.read_excel(ExcelName,sheet_name=Sheet,header=None,usecols=SRange,nrows=SheetRange[i][1])
            current_sheet = current_sheet.iloc[SheetRange[i][0]-1:SheetRange[i][1]]
            Cols = current_sheet.shape[1]
            #print(current_sheet,Cols)
            #write 
            fout.write(Sheet+':\n')
            fout.write('[\n')
            for colid in range(Cols):
                Sheet_1_col = current_sheet.iloc[:,colid]
                Sheet_Drop_NAN = Sheet_1_col.dropna()
                Sheet_list = Sheet_Drop_NAN.values.tolist()
                fout.write(str(Sheet_list).replace("'",""))
                fout.write(',\n')
                #print(str(Sheet_list).replace("'",""))
            fout.write(']\n\n\n')
    return
        


if __name__== '__main__':
    print('Parameter file name:',sys.argv[1])
    with open(sys.argv[1], 'r') as fin:
        params = json.load(fin)
    print('Check the parameters:')
    print(params)
    ExcelToTxt(params)