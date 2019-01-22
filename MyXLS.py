#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import os

class XLSbase(object):
    FieldsReferences = {'ALL': {
        u'LINE_1_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        u'LINE_2_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        u'STATUS_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        u'SELECT_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        u'ACKN_TYPE': [u'ALARM_ACK_DF']
        },
        'ITEM_DF': {
            u'SECTION_PATH': [u'SECTION_DF', u'OBJECT_DF'],
            u'ALARM_GROUP': [u'ALARM_GROUP_DF'],
            u'COL_GROUP': [u'ALARM_FU_DF'],
            u'FO_GROUP': [u'ALARM_FO_DF'],
            u'ITEM_STAT_1': [u'STATUS_DF'],
            u'ITEM_STAT_2': [u'STATUS_DF'],
            u'ITEM_STAT_3': [u'STATUS_DF'],
            u'ITEM_STAT_4': [u'STATUS_DF'],
            u'ITEM_STAT_5': [u'STATUS_DF'],
            u'ITEM_STAT_6': [u'STATUS_DF'],
            u'OPC_AE_STATION_NAME': [u'OPCAEC_STATION_DF'],
            u'OPC_EVENT_SOURCE': [],
            u'OPC_EVENT_SOURCE_NAME': [],
            u'INSTALL': [u'INSTALL_DF'],
            u'UNIT': [u'UNIT_DF'],
            u'POINT_NAME': [u'ABCIP_POINT_DF', u'ABPLC5_POINT_DF', u'ACGATEWAY_POINT_DF',
                            u'BKHFBK8100_POINT_DF', u'BRISTOLBCK_POINT_DF', u'DAQSTATION_POINT_DF',
                            u'DNP3_POINT_DF', u'DTS_POINT_DF', u'FAM3_POINT_DF',
                            u'FISHERROC_POINT_DF', u'HEXREPEATER_POINT_DF', u'HOSTHOST_POINT_DF',
                            u'IEC101_POINT_DF', u'IEC102_POINT_DF', u'IEC103_POINT_DF',
                            u'IEC104_POINT_DF', u'IEC61850_POINT_DF', u'MELSEC_POINT_DF',
                            u'MODBUS_POINT_DF', u'OPCDAC_POINT_DF', u'OPCUAC_POINT_DF',
                            u'PROSAFECOM_POINT_DF', u'PROSAFEPLC_POINT_DF', u'SAPIS7_POINT_DF',
                            u'SIEMENS3964_POINT_DF', u'STARDOMFCX_POINT_DF', u'STXBACHMANN_POINT_DF',
                            u'TIE8705_POINT_DF', u'VNET_POINT_DF'],
            u'STATION': [u'ACGATEWAY_STATION_DF', u'BKHFBK8100_STATION_DF', u'DAQSTATION_STATION_DF',
                         u'DTS_STATION_DF', u'FAM3_STATION_DF', u'HOSTHOST_STATION_DF',
                         u'MODBUS_STATION_DF', u'OPCAEC_STATION_DF', u'OPCDAC_STATION_DF',
                         u'OPCUAC_STATION_DF', u'OSIPI_STATION_DF', u'PROSAFECOM_STATION_DF',
                         u'PROSAFEPLC_STATION_DF', u'STARDOMFCX_STATION_DF', u'STXBACHMANN_STATION_DF'],
            u'POINT useless': [u'ABCIP_POINT_DF', u'ABPLC5_POINT_DF', u'ACGATEWAY_POINT_DF',
                       u'BKHFBK8100_POINT_DF', u'BRISTOLBCK_POINT_DF', u'DAQSTATION_POINT_DF',
                       u'DNP3_POINT_DF', u'DTS_POINT_DF', u'FAM3_POINT_DF',
                       u'FISHERROC_POINT_DF', u'HEXREPEATER_POINT_DF', u'HOSTHOST_POINT_DF',
                       u'IEC101_POINT_DF', u'IEC102_POINT_DF', u'IEC103_POINT_DF',
                       u'IEC104_POINT_DF', u'IEC61850_POINT_DF', u'MELSEC_POINT_DF',
                       u'MODBUS_POINT_DF', u'OPCDAC_POINT_DF', u'OPCUAC_POINT_DF',
                       u'PROSAFECOM_POINT_DF', u'PROSAFEPLC_POINT_DF', u'SAPIS7_POINT_DF',
                       u'SIEMENS3964_POINT_DF', u'STARDOMFCX_POINT_DF', u'STXBACHMANN_POINT_DF',
                       u'TIE8705_POINT_DF', u'VNET_POINT_DF'],
            u'AOI_1': [u'ALARM_AOI_DF'],
            u'AOI_2': [u'ALARM_AOI_DF'],
            u'AOI_3': [u'ALARM_AOI_DF'],
            u'AOI_4': [u'ALARM_AOI_DF'],
            u'AOI_5': [u'ALARM_AOI_DF'],
            u'AOI_6': [u'ALARM_AOI_DF'],
            u'AOI_7': [u'ALARM_AOI_DF'],
            u'AOI_8': [u'ALARM_AOI_DF'],
            u'AOI_9': [u'ALARM_AOI_DF'],
            u'AOI_10': [u'ALARM_AOI_DF'],
            u'AOI_11': [u'ALARM_AOI_DF'],
            u'AOI_12': [u'ALARM_AOI_DF'],
            u'AOI_13': [u'ALARM_AOI_DF'],
            u'AOI_14': [u'ALARM_AOI_DF'],
            u'AOI_15': [u'ALARM_AOI_DF'],
            u'AOI_16': [u'ALARM_AOI_DF'],
        },

        'SUB_ITEM_DF': {
            u'SECTION_PATH': [u'SECTION_DF'],
            u'ALARM_GROUP': [u'ALARM_GROUP_DF'],
            u'COL_GROUP': [u'ALARM_FU_DF'],
            u'FO_GROUP': [u'ALARM_FO_DF'],
            u'ITEM_STAT_1': [u'STATUS_DF'],
            u'ITEM_STAT_2': [u'STATUS_DF'],
            u'ITEM_STAT_3': [u'STATUS_DF'],
            u'ITEM_STAT_4': [u'STATUS_DF'],
            u'ITEM_STAT_5': [u'STATUS_DF'],
            u'ITEM_STAT_6': [u'STATUS_DF'],
            u'INSTALL': [u'INSTALL_DF'],
            u'UNIT': [u'UNIT_DF'],
            u'TAG': [u'ITEM_DF'],
            u'AOI_1': [u'ALARM_AOI_DF'],
            u'AOI_2': [u'ALARM_AOI_DF'],
            u'AOI_3': [u'ALARM_AOI_DF'],
            u'AOI_4': [u'ALARM_AOI_DF'],
            u'AOI_5': [u'ALARM_AOI_DF'],
            u'AOI_6': [u'ALARM_AOI_DF'],
            u'AOI_7': [u'ALARM_AOI_DF'],
            u'AOI_8': [u'ALARM_AOI_DF'],
            u'AOI_9': [u'ALARM_AOI_DF'],
            u'AOI_10': [u'ALARM_AOI_DF'],
            u'AOI_11': [u'ALARM_AOI_DF'],
            u'AOI_12': [u'ALARM_AOI_DF'],
            u'AOI_13': [u'ALARM_AOI_DF'],
            u'AOI_14': [u'ALARM_AOI_DF'],
            u'AOI_15': [u'ALARM_AOI_DF'],
            u'AOI_16': [u'ALARM_AOI_DF'],
        },
        'OBJECT_DF': {
            u'CLASS': [u'CLASS_DF']
        },
        'ITEM_HIS_DF': {
            u'ITEM_NAME': [u'ITEM_DF', u'SUB_ITEM_DF'],
            u'GROUP_NAME': [u'HIS_GROUP_DF']
        },
        'SEQUENCE_df': {
            u'NAME': [u'ITEM_DF', u'SUB_ITEM_DF']
        },
        'REPORT_DF': {
            u'DEST_NAME': [u'PRINTER_DEST_DF'],
            u'TRIGGER_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        },
        'PRINTER_DEST_DF': {
            u'PRIMARY_DEVICE': [u'PRINTER_DEV_DF'],
            u'SECONDARY_DEVICE': [u'PRINTER_DEV_DF'],
        },
        'USER_DF': {
            u'AUTH_GROUP': [u'AUTH_GROUP_DF'],
            u'ASA': [u'ALARM_ASA_DF'],
            u'ALARM_SHELF_GROUP': [u'ALARM_SHELF_GROUP_DF'],
            u'RPT_DEST': [u'PRINTER_DEST_DF'],
        },
        'ALARM_FU_DF': {
            u'RESET_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
            u'ACKN_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
        },
        'ALARM_SHELF_GROUP_DF': {
            u'SHELF_TYPE_1': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_2': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_3': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_4': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_5': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_6': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_7': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_8': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_9': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_10': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_11': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_12': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_13': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_14': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_15': [u'ALARM_SHELF_DF'],
            u'SHELF_TYPE_16': [u'ALARM_SHELF_DF'],
        },
        'ALARM_NOT_USR': {
            u'USER_NAME': [u'USER_DF'],
        },
        'ALARM_NOT_DEST': {
            u'DEST_ASA': [u'ALARM_ASA_DF'],
            u'CALLBACK_ITEM': [u'ITEM_DF', u'SUB_ITEM_DF'],
            u'NOT_USR_1': [u'ALARM_NOT_USR'],
            u'NOT_USR_2': [u'ALARM_NOT_USR'],
            u'NOT_USR_3': [u'ALARM_NOT_USR'],
            u'NOT_USR_4': [u'ALARM_NOT_USR'],
            u'NOT_USR_5': [u'ALARM_NOT_USR'],
            u'NOT_USR_6': [u'ALARM_NOT_USR'],
            u'NOT_USR_7': [u'ALARM_NOT_USR'],
            u'NOT_USR_8': [u'ALARM_NOT_USR'],
            u'NOT_USR_9': [u'ALARM_NOT_USR'],
            u'NOT_USR_10': [u'ALARM_NOT_USR'],
            u'NOT_USR_11': [u'ALARM_NOT_USR'],
            u'NOT_USR_12': [u'ALARM_NOT_USR'],
            u'NOT_USR_13': [u'ALARM_NOT_USR'],
            u'NOT_USR_14': [u'ALARM_NOT_USR'],
            u'NOT_USR_15': [u'ALARM_NOT_USR'],
            u'NOT_USR_16': [u'ALARM_NOT_USR'],
            u'NOT_USR_17': [u'ALARM_NOT_USR'],
            u'NOT_USR_18': [u'ALARM_NOT_USR'],
            u'NOT_USR_19': [u'ALARM_NOT_USR'],
            u'NOT_USR_20': [u'ALARM_NOT_USR'],
            u'NOT_USR_21': [u'ALARM_NOT_USR'],
            u'NOT_USR_22': [u'ALARM_NOT_USR'],
            u'NOT_USR_23': [u'ALARM_NOT_USR'],
            u'NOT_USR_24': [u'ALARM_NOT_USR'],
            u'NOT_USR_25': [u'ALARM_NOT_USR'],
            u'NOT_USR_26': [u'ALARM_NOT_USR'],
            u'NOT_USR_27': [u'ALARM_NOT_USR'],
            u'NOT_USR_28': [u'ALARM_NOT_USR'],
            u'NOT_USR_29': [u'ALARM_NOT_USR'],
            u'NOT_USR_30': [u'ALARM_NOT_USR'],
            u'NOT_USR_31': [u'ALARM_NOT_USR'],
            u'NOT_USR_32': [u'ALARM_NOT_USR']
        },
        'ALARM_DISPLAY_DF': {
            u'USER_NAME': [u'USER_DF'],
            u'ITEM': [u'ITEM_DF', u'SUB_ITEM_DF']
        },
        'ABCIP_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'ABCIP_LINE_DF']
        },
        'ABPLC5_POINT_DF': {
            u'SCAN_TYPE': [u'ABPLC5_SCAN_TYPE_DF'],
            u'STATION': [u'ABPLC5_LINE_DF']
        },
        'ACGATEWAY_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'ACGATEWAY_STATION_DF']
        },
        'ACGATEWAY_STATION_DF':{
            u'LINE':[u'ACGATEWAY_LINE_DF']
        },
        'BKHFBK8100_POINT_DF': {
            u'SCAN_TYPE': [u'BKHFBK8100_SCAN_TYPE_DF'],
            u'STATION': [u'BKHFBK8100_STATION_DF']
        },
        'BKHFBK8100_STATION_DF': {
            u'LINE': [u'BKHFBK8100_LINE_DF']
        },
        'BRISTOLBCK_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'BRISTOLBCK_LINE_DF']
        },
        'DAQSTATION_POINT_DF': {
            u'SCAN_TYPE': [u'DAQSTATION_SCAN_TYPE_DF'],
            u'STATION': [u'DAQSTATION_STATION_DF']
        },
        'DAQSTATION_STATION_DF': {
            u'LINE': [u'DAQSTATION_LINE_DF']
        },
        'DNP3_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'DNP3_LINE_DF']
        },
        'DTS_POINT_DF': {
            u'SCAN_TYPE': [u'DTS_SCAN_TYPE_DF'],
            u'STATION': [u'DTS_STATION_DF']
        },
        'DTS_STATION_DF': {
            u'LINE': [u'DTS_LINE_DF']
        },
        'FAM3_POINT_DF': {
            u'SCAN_TYPE': [u'FAM3_SCAN_TYPE_DF'],
            u'STATION': [u'FAM3_STATION_DF']
        },
        'FAM3_STATION_DF': {
            u'LINE': [u'FAM3_LINE_DF']
        },
        'FISHERROC_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'FISHERROC_LINE_DF']
        },
        'HEXREPEATER_POINT_DF': {
            u'STATION': [u'HEXREPEATER_LINE_DF']
        },
        'HOSTHOST_POINT_DF': {
            u'STATION': [u'HOSTHOST_STATION_DF']
        },
        'HOSTHOST_STATION_DF': {
            u'LINE': [u'HOSTHOST_LINE_DF']
        },
        'IEC101_POINT_DF': {
            u'STATION': [u'IEC101_LINE_DF']
        },
        'IEC102_POINT_DF': {
            u'SCAN_TYPE': [u'IEC102_SCAN_TYPE_DF'],
            u'STATION': [u'IEC102_LINE_DF']
        },
        'IEC103_POINT_DF': {
            u'STATION': [u'IEC103_LINE_DF']
        },
        'IEC104_POINT_DF': {
            u'STATION': [u'IEC104_LINE_DF']
        },
        'IEC61850_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'IEC61850_LINE_DF']
        },
        'MELSEC_POINT_DF': {
            u'SCAN_TYPE': [u'MELSEC_SCAN_TYPE_DF'],
            u'STATION': [u'MELSEC_LINE_DF']
        },
        'MODBUS_POINT_DF': {
            u'SCAN_TYPE': [u'MODBUS_SCAN_TYPE_DF'],
            u'STATION': [u'MODBUS_STATION_DF']
        },
        'MODBUS_STATION_DF': {
            u'LINE': [u'MODBUS_LINE_DF']
        },

        'OPCAEC_STATION_DF': {
            u'LINE': [u'OPCAEC_LINE_DF']
        },
        'OPCDAC_POINT_DF': {
            u'STATION': [u'OPCDAC_STATION_DF'],
            u'OPC_GROUP': [u'OPCDAC_GROUP_DF']
        },
        'OPCDAC_STATION_DF': {
            u'LINE': [u'OPCDAC_LINE_DF']
        },
        'OPCUAC_POINT_DF': {
            u'STATION': [u'OPCUAC_STATION_DF'],
            u'SUBSCRIPTION': [u'OPCUAC_STATION_DF']
        },
        'OPCUAC_SUBSCRIPTION_DF': {
            u'STATION_NAME': [u'OPCUAC_STATION_DF']
        },
        'OPCUAC_STATION_DF': {
            u'LINE': [u'OPCUAC_LINE_DF']
        },
        'PROSAFECOM_POINT_DF': {
            u'SCAN_TYPE': [u'PROSAFECOM_SCAN_TYPE_DF'],
            u'STATION': [u'PROSAFECOM_STATION_DF'],
        },
        'PROSAFECOM_STATION_DF': {
            u'LINE': [u'PROSAFECOM_LINE_DF']
        },
        'PROSAFEPLC_POINT_DF': {
            u'SCAN_TYPE': [u''],
            u'STATION': [u'PROSAFEPLC_STATION_DF'],
        },
        'PROSAFEPLC_STATION_DF': {
            u'LINE': [u'PROSAFEPLC_LINE_DF']
        },
        'SAPIS7_POINT_DF': {
            u'SCAN_TYPE': [u'SAPIS7_SCAN_TYPE_DF'],
            u'STATION': [u'SAPIS7_LINE_DF'],
        },
        'SIEMENS3964_POINT_DF': {
            u'SCAN_TYPE': [u'SIEMENS3964_SCAN_TYPE_DF'],
            u'STATION': [u'SIEMENS3964_LINE_DF'],
        },
        'STARDOMFCX_POINT_DF': {
            u'STATION': [u'STARDOMFCX_STATION_DF'],
        },
        'STARDOMFCX_STATION_DF': {
            u'FF_MESSAGE_ITEM': [u'ITEM_DF'],
            u'GENERAL_MESSAGE_ITEM': [u'ITEM_DF'],
            u'LINE': [u'STARDOMFCX_LINE_DF']
        },
        'STXBACHMANN_POINT_DF': {
            u'SCAN_TYPE': [u'STXBACHMANN_SCAN_TYPE_DF'],
            u'STATION': [u'STXBACHMANN_STATION_DF'],
        },
        'STXBACHMANN_STATION_DF': {
            u'LINE': [u'STXBACHMANN_LINE_DF']
        },
        'TIE8705_POINT_DF': {
            u'STATION': [u'TIE8705_LINE_DF'],
        },
        'VNET_POINT_DF': {
            u'STATION': [u'VNET_LINE_DF'],
        }
    }

    def __init__(self, folderpath):
        self.DataSets = []
        self.DataSetsNames = []
        self.GetDataSetsFromFiles(folderpath)
        #self.CheckDataSetsItems()
        self.CheckDataSetsReferences()
        self.Tidy()

    def GetDataSetsFromFiles(self, folderpath):
        u'''Получаем датасэты из файлов'''
        DataSets = []
        for xls in os.listdir(folderpath):
            if xls.split('.')[-1] == 'xls':
                self.DataSets.append(DataSet(folderpath, xls))
                self.DataSetsNames.append(self.DataSets[-1].Name.upper())

    def CheckDataSetsReferences(self):
        #Пробегаем по всем датасетам
        for dataset in self.DataSets:
            #Если в дата сете есть элементы
            if dataset.ItemsCount > 0:
                #dataset.GetItemsNames()
                dsName = dataset.Name.upper()
                #Проверяем есть ли датасэт в перечне датасэтов с ссылками на другие датасэты
                if dsName in self.FieldsReferences:
                    FieldsReferences = self.FieldsReferences['ALL'].copy()
                    FieldsReferences.update(self.FieldsReferences[dsName])
                    #Пробегаем по полям датасэта
                    for field in dataset.Fields:
                        foundfields = []
                        lostfields = []
                        #Проверяем поля датасэта на наличие их в списке полей с ссылками, с условием что в этих полях есть что-либо вообще
                        if field.upper() in FieldsReferences and field not in dataset.EmptyFields:
                            #Теперь проверяем те датасэты, на которые мы ссылаемся из поля основного датасэта
                            for otherDataSet in FieldsReferences[field.upper()]:
                                ODS = self.GetDatsSebByName(otherDataSet)
                                #Проверяем, есть ли вообще в выгрузке датасэт, на который мы ссылаемся и если есть проверяем есть ли в нём элементы
                                if otherDataSet in self.DataSetsNames and ODS.ItemsCount > 0:
                                    #print dataset.Name, ODS.Name, '=============================================================='
                                    #print ODS.ItemsNames
                                    #Пробегаемся по всем элементам основного датасэта
                                    for item in dataset.Items:
                                        #Проверяем, есть ли данный элемент в перечне найденных полей
                                        if item not in foundfields:
                                            if item.__dict__[field].upper() in ODS.ItemsNames:
                                                #print field, item.__dict__[field], ODS.Name
                                                foundfields.append(item)
                                                if item in lostfields:
                                                    lostfields.remove(item)
                                                if ODS not in dataset.FirstOrder:
                                                    dataset.FirstOrder.append(ODS)
                                            elif item.__dict__[field] != '' and item.__dict__[field] != ':':
                                                #print field, item.__dict__[field], ODS.Name, '****************************'
                                                lostfields.append(item)
                            for item in lostfields:
                                print u'В датасете', dataset.Name, u'для элемента', item.Name, u'для поля', field, u'с значением', item.__dict__[field], u'не найдены соответсвующие данные в датасете', FieldsReferences[field.upper()]

    def GetListExcludedReferenFields(self, dataset, NameRefDS):
        ListOfFields = []
        dsName = dataset.Name.upper()
        FieldsReferences = self.FieldsReferences['ALL'].copy()
        FieldsReferences.update(self.FieldsReferences[dsName])
        for field in dataset.Fields:
            # Проверяем поля датасэта на наличие их в списке полей с ссылками, с условием что в этих полях есть что-либо вообще
            if field.upper() in FieldsReferences and field not in dataset.EmptyFields:
                for otherDataSet in FieldsReferences[field.upper()]:
                    if otherDataSet == NameRefDS:
                        ListOfFields.append(field.upper())
        return ListOfFields

    def GetDatsSebByName(self, Name):
        for dataset in self.DataSets:
            if dataset.Name.upper() == Name.upper():
                return dataset
        else:
            return None

    def Tidy(self):
        obj = self.GetDatsSebByName(u'OBJECT_DF')
        itm = self.GetDatsSebByName(u'ITEM_DF')
        objnames = []
        if obj and itm:
            for item in obj.Items:
                objnames.append(item.Name)
            tempitmItems = itm.Items [:]
            for item in tempitmItems:
                if '.' in item.Name:
                    section = '.'.join(item.Name.split('.')[:-1])
                else:
                    section = ''
                if section in objnames:
                    #print item.Name
                    itm.Items.remove(item)
                    if item.__dict__.get('Point_name') and item.Point_name != ':':
                        for dspointname in self.FieldsReferences['ITEM_DF']['POINT_NAME']:
                            breakflag = False
                            pnt = self.GetDatsSebByName(dspointname)
                            if pnt:
                                temppntItems = pnt.Items[:]
                                for point in temppntItems:
                                    if point.Name == item.Point_name:
                                        breakflag = True
                                        pnt.Items.remove(point)
                                        pnt.ItemsCount = len(pnt.Items)
                                        break
                            if breakflag:
                                break
            itm.ItemsCount = len(itm.Items)
        pointnames = []
        if itm:
            for item in itm.Items:
                pointnames.append(item.Point_name)
            for dspointname in self.FieldsReferences['ITEM_DF']['POINT_NAME']:
                pnt = self.GetDatsSebByName(dspointname)
                if pnt:
                    temppntItems = pnt.Items[:]
                    for point in temppntItems:
                        if point.Name not in pointnames:
                            pnt.Items.remove(point)
                            pnt.ItemsCount = len(pnt.Items)

    def SayWhatYouHave(self):
        for k in self.DataSets:
            print k.Name, u'FieldsCount', k.FieldsCount, u'Itemcount', k.ItemsCount

class DataSet(object):
    u'''Дата сэт'''
    def __init__(self, folderpath, filename):
        u'''Создаем дата сет с его аттрибутами'''
        self.Name = filename.split('.')[0]
        self.Items = []
        self.ItemsCount = 0
        self.ItemsNames = []
        self.Fields = []
        self.EmptyFields = []
        self.FieldsCount = 0
        self.FirstOrder = []

        rb = xlrd.open_workbook(folderpath + u'\\' + filename)
        sheet = rb.sheet_by_index(0)
        ncols = sheet.ncols
        # Получаем поля и их количество
        for col in range(ncols):
            cell = sheet.cell_value(0, col)
            if type(cell) != unicode:
                print u'Unicode Alarm'
            self.Fields.append(cell.capitalize())
        else:
            self.FieldsCount = len(self.Fields)
        # Получаем элементы дата сэта
        for row in range(1, sheet.nrows):
            ItemFields = []
            for col in range(ncols):
                cell = sheet.cell_value(row, col)
                if cell != '' and type(cell) != unicode:
                    print u'Unicode Alarm',row, col, 'value', cell, 'file', filename
                    cell = unicode(cell)
                if cell == None:
                    cell = u''
                if cell[:9] == u'File_@"@_':
                    cell = self.GetDataFromTxtFile(folderpath, cell[9:])
                ItemFields.append(cell)
            if len(ItemFields) == len(self.Fields):
                self.Items.append(ItemOfDataSet(self.Fields, ItemFields))
            else:
                print u'Проблемы с датасэтом'
        else:
            self.ItemsCount = len(self.Items)
        #Проверяем неиспользуемые (во всех элементах пустое) поля
        self.CheckDataSetsItems()
        self.GetEmptyFields()
        self.GetItemsNames()

    def CheckDataSetsItems(self):
        if self.Name.upper() == u'ITEM_DF':
            for item in self.Items:
                #сначала сбор и анализ данных подлежащих проверке
                if item.__dict__.get('Name') and item.Name != '':
                    name = item.Name
                    if '.' in item.Name:
                        partsofname = item.Name.split('.')
                        section = '.'.join(partsofname[:-1])
                        tag = partsofname[-1]
                        if len(partsofname) > 2:
                            install = partsofname[0]
                            unit = partsofname[1]
                        else:
                            install = partsofname[0]
                            unit = ''
                    else:
                        tag = item.Name
                        section = ''
                        install = ''
                        unit = ''
                elif item.__dict__.get('Tag') and item.Tag != '':
                    tag = item.Tag
                    if item.__dict__.get('Section_path') and item.Section_path != '':
                        section = item.Section_path
                        name = section + '.' + tag
                        if '.' in item.Section_path:
                            partsofpath = item.Section_path.split('.')
                            install = partsofpath[0]
                            unit = partsofpath[1]
                        else:
                            install = section
                            unit = ''
                else:
                    print u'Критическая ошибка в датасете ITEM_DF нет ни имени айтема Name, ни имени тега Tag (нужно хотя бы что-то одно'
                if item.__dict__.get('Point_name') and item.Point_name != '' and item.Point_name != ':':
                    pointname = item.Point_name
                    partsofpoint = item.Point_name.split(':')
                    station = partsofpoint[0]
                    point = partsofpoint[1]
                elif item.__dict__.get('Station') and item.Station != '' and item.__dict__.get('Point') and item.Point != '':
                    station = item.Station
                    point = item.Point
                    pointname = station + ':' + point
                else:
                    pointname = ':'
                    station = ''
                    point = ''
                #Теперь внесение данных
                item.Name = name
                if u'Name' not in self.Fields:
                    self.Fields.insert(0, u'Name')
                    self.FieldsCount += 1
                item.Point_name = pointname
                if u'Point_name' not in self.Fields:
                    self.Fields.insert(1, u'Point_name')  #it was u'Name'
                    self.FieldsCount += 1
                if item.__dict__.get('Section_path'):
                    item.Section_path = section
                if item.__dict__.get('Tag'):
                    item.Tag = tag
                if item.__dict__.get('Install'):
                    item.Install = install
                if item.__dict__.get('Unit'):
                    item.Unit = unit
                if item.__dict__.get('Station'):
                    item.Station = station
                if item.__dict__.get('Point'):
                    item.Point = point
        elif self.Name.upper() in [u'ABCIP_POINT_DF', u'ABPLC5_POINT_DF', u'ACGATEWAY_POINT_DF',
                                   u'BKHFBK8100_POINT_DF', u'BRISTOLBCK_POINT_DF', u'DAQSTATION_POINT_DF',
                                   u'DNP3_POINT_DF', u'DTS_POINT_DF', u'FAM3_POINT_DF',
                                   u'FISHERROC_POINT_DF', u'HEXREPEATER_POINT_DF', u'HOSTHOST_POINT_DF',
                                   u'IEC101_POINT_DF', u'IEC102_POINT_DF', u'IEC103_POINT_DF',
                                   u'IEC104_POINT_DF', u'IEC61850_POINT_DF', u'MELSEC_POINT_DF',
                                   u'MODBUS_POINT_DF', u'OPCDAC_POINT_DF', u'OPCUAC_POINT_DF',
                                   u'PROSAFECOM_POINT_DF', u'PROSAFEPLC_POINT_DF', u'SAPIS7_POINT_DF',
                                   u'SIEMENS3964_POINT_DF', u'STARDOMFCX_POINT_DF', u'STXBACHMANN_POINT_DF',
                                   u'TIE8705_POINT_DF', u'VNET_POINT_DF']:
            for item in self.Items:
                if item.__dict__.get('Name') and item.Name != '':
                    name = item.Name
                    partsofpoint = item.Name.split(':')
                    station = partsofpoint[0]
                    point = partsofpoint[1]
                elif item.__dict__.get('Station') and item.Station != '' and item.__dict__.get('Point') and item.Point != '':
                    station = item.Station
                    point = item.Point
                    name = station + ':' + point
                else:
                    print u'Критическая ошибка в датасете ' + self.Name + u' нет данных о point'
                item.Name = name
                if u'Name' not in self.Fields:
                    self.Fields.insert(0, u'Name')
                    self.FieldsCount += 1
                if item.__dict__.get('Station'):
                    item.Station = station
                if item.__dict__.get('Point'):
                    item.Point = point
        elif self.Name.upper() == u'ITEM_HIS_DF':
            for item in self.Items:
                #сначала сбор и анализ данных подлежащих проверке
                if item.__dict__.get('Name') and item.Name != '':
                    partsofhis = item.Name.split(':')
                    item.Group_name = partsofhis[0]
                    item.Item_name = partsofhis[1]
                elif item.__dict__.get('Group_name') and item.Group_name != '' and item.__dict__.get('Item_name') and item.Item_name != '':
                    item.Name = item.Group_name + ':' + item.Item_name
                    if u'Name' not in self.Fields:
                        self.Fields.insert(0, u'Name')
                        self.FieldsCount += 1
                else:
                    print u'Критическая ошибка в датасете ' + self.Name
        elif self.Name.upper() == u'ALARM_DISPLAY_DF':
            for item in self.Items:
                #сначала сбор и анализ данных подлежащих проверке
                if item.__dict__.get('Name') and item.Name != '':
                    partsofalrd = item.Name.split(':')
                    item.Item = partsofalrd[0]
                    item.Display_name = partsofalrd[2]
                elif item.__dict__.get('Item') and item.Item != '' and item.__dict__.get('Display_name') and item.Display_name != '':
                    item.Name = item.Item + '::' + item.Display_name
                    if u'Name' not in self.Fields:
                        self.Fields.insert(0, u'Name')
                        self.FieldsCount += 1
                else:
                    print u'Критическая ошибка в датасете ' + self.Name

    def GetEmptyFields(self):
        self.EmptyFields = self.Fields[:]
        for field in self.Fields:
            for item in self.Items:
                if item.__dict__[field] != '' or item.__dict__[field] != ':':
                    self.EmptyFields.remove(field)
                    break
        #print self.Name, self.Fields, self.EmptyFields

    def GetItemsNames(self):
        #Вызываем после CheckDataSetsItems, потому что в исходном виде может отсутствовать атрибут Name
        self.ItemsNames = []
        for item in self.Items:
            self.ItemsNames.append(item.Name.upper())

    def AddElementInOrder(self, NameOfDataSet):
        #Добавить датасэт в список датасэтов, которые должны быть загружены раньше чем, этот
        if NameOfDataSet not in self.FirstOrder:
            self.FirstOrder.append(NameOfDataSet)

    def GetDataFromTxtFile(self, folderpath, filename):
        result = []
        for txt in os.listdir(folderpath):
            if txt.split('.')[-1] == 'txt' and txt.split('.')[0] == filename:
                file = open(folderpath + u'\\' + txt, 'r')
                for line in file:
                    if type(line) == str:
                        line = line.decode('cp1251')
                    result.append(line)
                break
        return result

class ItemOfDataSet(object):
    u'''Элемент датасэта'''
    #FoundFieldsODS = []
    def __init__(self, Fields, Data):
        for i in range(len(Fields)):
            self.__dict__[Fields[i]] = Data[i]

    def SayWhatYouHave(self):
        for attr, value in self.__dict__.items():
            print attr, value
        print u'*******************************************'

if __name__ == "__main__":
    b = XLSbase('C:\\Share\\TEST ASU IS\\RESULT')
    #b.SayWhatYouHave()