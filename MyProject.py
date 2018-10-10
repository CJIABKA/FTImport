#!/usr/bin/env python
# -*- coding: utf-8 -*-
from MyXLS import XLSbase
#from MyDevices_lib import Device
from MyFastTools_lib import FastTools
import os
import xlwt


class Project(object):
    u'''Проект!! Здесь всё тело программы'''
    FieldsCountInRow = 6
    ExcludeFields = {u'ITEM_DF': [u'NSID', u'PARENT_NSID', u'ID_NUMBER' , u'FRONT_END_NODE' , u'DISTR_TYPE'],
                     u'OBJECT_DF': [u'NSID', u'PARENT_NSID', u'NODE_NAME'],
                     u'ALARM_AOI_DF': [u'NUMBER'],
                     u'ALARM_FU_DF': [u'NODE'],
                     u'SECTION_DF': [u'NSID', u'PARENT_NSID'],
                     u'PROSAFEPLC_LINE_DF': [u'EQUIPMENT_NODE'],
                     u'PROSAFECOM_LINE_DF': [u'EQUIPMENT_NODE'],
                     u'STARDOMFCX_LINE_DF': [u'EQUIPMENT_NODE']
                     }

    AfterItemDFupdate = [u'ACGATEWAY_STATION_DF', u'BKHFBK8100_STATION_DF', u'DAQSTATION_STATION_DF',
                         u'DTS_STATION_DF', u'FAM3_STATION_DF', u'HOSTHOST_STATION_DF',
                         u'MODBUS_STATION_DF', u'OPCAEC_STATION_DF', u'OPCDAC_STATION_DF',
                         u'OPCUAC_STATION_DF', u'OSIPI_STATION_DF', u'PROSAFECOM_STATION_DF',
                         u'PROSAFEPLC_STATION_DF', u'STARDOMFCX_STATION_DF', u'STXBACHMANN_STATION_DF',
                         u'ALARM_FU_DF']


    def __init__(self, BasePath, ProjectPath):
        self.Base = XLSbase(BasePath)
        self.ProjectPath = ProjectPath
        self.__CheckPath()
        self.Order = []
        self.OrderUpdate = []
        self.SortDataSetsByOrder()

    def __CheckPath(self, Path=[]):
        u'''Проверяет путь указанный как "путь проекта" и если данная папка не создана - создаёт'''
        folders = self.ProjectPath.split('\\')
        folders = folders + Path
        path = ''
        for folder in folders:
            path = path + folder if path == '' else path + '\\' + folder
            if not os.path.exists(path):
                os.mkdir(path)

    def MakeXLS(self, Path, Table, FileName):
        u'''Запись значений переданных в двумерном массиве в таблицу по указанному пути в папке проекта'''
        try:
            self.__CheckPath(Path)
            wb = xlwt.Workbook()
            sheets = Table.keys()
            sheets.sort(key=lambda x: x.split('|')[0])
            for sheet in sheets:
                ws = wb.add_sheet(sheet.split('|')[1])
                count_row = 0
                for row in Table[sheet]:
                    count_column = 0
                    for column in row:
                        ws.write(count_row, count_column, column)
                        count_column += 1
                    count_row += 1
            Path.append(FileName)
            wb.save(self.ProjectPath + '\\' + '\\'.join(Path))
        except Exception as ex:
            print u'Проблемы с записью файла Excel', Path, row, ex

    def MakeTextFile(self, Path, Rows, FileName):
        u'''Запись строк переданных в списке в текстовый файл по указанному пути в папке проекта'''
        row = ''
        try:
            self.__CheckPath(Path)
            Path.append(FileName)
            f = open(self.ProjectPath + '\\' + '\\'.join(Path), 'w')
            for row in Rows:
                if row[-1] == '\n':
                    f.write(row)
                else:
                    f.write(row + '\n')
            f.close()
        except Exception as ex:
            print u'Проблемы с записью файла txt', Path, row, ex

    def SortDataSetsByOrder(self):
        DataSets = []
        for DataSet in self.Base.DataSets:
            if DataSet.ItemsCount > 0:
                DataSets.append(DataSet)
        previouslen = None
        loopbrakecounter = 0
        while len(DataSets) > 0:
            if previouslen == len(DataSets):
                loopbrakecounter += 1
                if loopbrakecounter > 10:
                    print u'Ошибка!! Бесконечный цикл во время сортировки датасэтов'
                    break
                for DataSet in DataSets:
                    if DataSet.Name in self.AfterItemDFupdate:
                        check = True
                        for ds in DataSet.FirstOrder:
                            if ds not in self.Order and ds.Name.upper() != u'ITEM_DF':
                                check = False
                        if check:
                            loopbrakecounter = 0
                            self.Order.append(DataSet)
                            self.OrderUpdate.append(DataSet)
                            DataSets.remove(DataSet)
                            break
            previouslen = len(DataSets)

            for DataSet in DataSets:
                if len(DataSet.FirstOrder) == 0:
                    self.Order.append(DataSet)
                    DataSets.remove(DataSet)
                    break
                check = True
                for ds in DataSet.FirstOrder:
                    if ds not in self.Order:
                        check = False
                if check:
                    self.Order.append(DataSet)
                    DataSets.remove(DataSet)
                    break
        obj = self.Base.GetDatsSebByName(u'OBJECT_DF')
        itm = self.Base.GetDatsSebByName(u'ITEM_DF')
        if obj and itm:
            self.Order.remove(obj)
            index = self.Order.index(itm)
            self.Order.insert(index+1, obj)

    def main(self):
        cmdstrings = ['echo off', 'attrib *.qlo -R', 'del *.qlo']
        cmdclassstrings = ['echo off', 'attrib *.qlo -R', 'del *.qlo']
        index = 1
        for DataSet in self.Order:
            dsstrings = ['@FIELDS\n']
            #print DataSet.Name
            datalist = []
            FieldsInRow = []
            if DataSet in self.OrderUpdate:
                ExcludeRefFields = self.Base.GetListExcludedReferenFields(DataSet, u'ITEM_DF')
            else:
                ExcludeRefFields = []
            for field in DataSet.Fields:
                ExcludeFields = self.ExcludeFields[DataSet.Name.upper()] if self.ExcludeFields.get(DataSet.Name.upper()) else []
                ExcludeFields = ExcludeFields + ExcludeRefFields
                if field.upper() not in ExcludeFields:
                    datalist.append(field.upper())
                if len(datalist) >= self.FieldsCountInRow:
                    dsstrings.append(','.join(datalist))
                    FieldsInRow.append(datalist)
                    datalist = []
            else:
                if len(datalist) > 0:
                    dsstrings.append(','.join(datalist))
                    FieldsInRow.append(datalist)
            dsstrings.append('@'+DataSet.Name.upper())

            for item in DataSet.Items:
                for Fields in FieldsInRow:
                    datalist = []
                    for field in Fields:
                        DataInField = item.__dict__[field.capitalize()]
                        if type(DataInField) == unicode:
                            DataInField = '"' + DataInField.encode('utf8') + '"'
                            datalist.append(DataInField)
                        elif type(DataInField) == str:
                            DataInField = '"' + DataInField + '"'
                            datalist.append(DataInField)
                        elif type(DataInField) == list:
                            DataInField = ''.join(DataInField)
                            if type(DataInField) == unicode:
                                DataInField = '@"\n' + DataInField.encode('utf8') + '"@'
                                datalist.append(DataInField)
                            elif type(DataInField) == str:
                                DataInField = '@"\n' + DataInField + '"@'
                                datalist.append(DataInField)
                        else:
                            datalist.append('')
                            print u'***********ОШИБКА В ТИПЕ ДАННЫХ ВО ВРЕМЯ ФОРМИРОВАНИЯ ФАЙЛА*********', DataInField, type(DataInField)
                    datalist.append('\\')
                    dsstrings.append(','.join(datalist))
                else:
                    dsstrings[-1] = dsstrings[-1][:-2]
                dsstrings.append('\n')


            datasetfilename = str(index).zfill(3) + '_I_' + DataSet.Name + '.qli'
            self.MakeTextFile([], dsstrings, datasetfilename)
            if DataSet.Name.upper() == u'CLASS_DF':
                cmdclassstrings.append('dssqld -i "' + datasetfilename + '" -l')
                cmdclassstrings.append('Echo Finished loading\nExit')
                self.MakeTextFile([], cmdclassstrings, u'__ImportFirstClassesToFT.cmd')
            cmdstrings.append('dssqld -i "' + datasetfilename + '" -l')
            index += 1
        for DataSet in self.OrderUpdate:
            dsstrings = ['@FIELDS\n']
            # print DataSet.Name
            datalist = []
            FieldsInRow = []
            for field in DataSet.Fields:
                ExcludeFields = self.ExcludeFields[DataSet.Name.upper()] if self.ExcludeFields.get(
                    DataSet.Name.upper()) else []
                if field.upper() not in ExcludeFields:
                    datalist.append(field.upper())
                if len(datalist) >= self.FieldsCountInRow:
                    dsstrings.append(','.join(datalist))
                    FieldsInRow.append(datalist)
                    datalist = []
            else:
                if len(datalist) > 0:
                    dsstrings.append(','.join(datalist))
                    FieldsInRow.append(datalist)
            dsstrings.append('@' + DataSet.Name.upper())

            for item in DataSet.Items:
                for Fields in FieldsInRow:
                    datalist = []
                    for field in Fields:
                        DataInField = item.__dict__[field.capitalize()]
                        if type(DataInField) == unicode:
                            DataInField = '"' + DataInField.encode('utf8') + '"'
                            datalist.append(DataInField)
                        elif type(DataInField) == str:
                            DataInField = '"' + DataInField + '"'
                            datalist.append(DataInField)
                        elif type(DataInField) == list:
                            DataInField = ''.join(DataInField)
                            if type(DataInField) == unicode:
                                DataInField = '@"\n' + DataInField.encode('utf8') + '"@'
                                datalist.append(DataInField)
                            elif type(DataInField) == str:
                                DataInField = '@"\n' + DataInField + '"@'
                                datalist.append(DataInField)
                        else:
                            datalist.append('')
                            print u'***********ОШИБКА В ТИПЕ ДАННЫХ ВО ВРЕМЯ ФОРМИРОВАНИЯ ФАЙЛА*********', DataInField, type(
                                DataInField)
                    datalist.append('\\')
                    dsstrings.append(','.join(datalist))
                else:
                    dsstrings[-1] = dsstrings[-1][:-2]
                dsstrings.append('\n')

            datasetfilename = str(index).zfill(3) + '_U_' + DataSet.Name + '.qli'
            self.MakeTextFile([], dsstrings, datasetfilename)
            cmdstrings.append('dssqld -m "' + datasetfilename + '" -l')
            index += 1
        cmdstrings.append('Echo Finished loading\nExit')
        self.MakeTextFile([], cmdstrings, u'_ImportToFT.cmd')