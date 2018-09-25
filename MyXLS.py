#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import os

class XLSbase(object):
    def __init__(self, folderpath):
        self.DataSets = self.GetDataSetsFromFiles(folderpath)

    def GetDataSetsFromFiles(self, folderpath):
        u'''Получаем датасэты из файлов'''
        DataSets = []
        for xls in os.listdir(folderpath):
            if xls[-4:] == '.xls':
                DataSet(folderpath,xls)
                rb = xlrd.open_workbook(folderpath + u'\\' + xls)
                sheet = rb.sheet_by_index(0)
                ncols = sheet.ncols
                #Fields
                for col in range(ncols):
                    cell = sheet.cell_value(0, col)
                    print cell
                #Data
                for row in range(sheet.nrows):
                    for col in range(ncols):
                        cell = sheet.cell_value(row, col)



    def SayWhatYouHave(self):
        for k in self.DataSets:
            print k.Name, u'FieldsCount', k.FieldsCount, u'Itemcount', k.ItemsCount

class DataSet(object):
    u'''Дата сэт'''
    def __init__(self, filepath, filename):
        u'''Создаем дата сет с его аттрибутами'''
        self.Name = ''
        self.Items = []
        self.ItemsCount = 0
        self.Fields = []
        self.FieldsCount = 0
        index = 0
        options = []
        for string in ListOfStrings:
            string = string.strip()
            if string != '' and string[0] == u'@' and string[-3:] == u'_DF':
                self.Name = string
                break
            elif string != '' and string[0] == u'@' and (',' not in string) and string != u'@FIELDS':
                options.append((string,index))
            index += 1
        else:
            self.Name = options[0][0]
            index = options[0][1]
            if len(options) > 1:
                print u'В датасэте возможны 2 имени'
        #Получаем поля и их количество
        for string in ListOfStrings[1:index]:
            string = string.strip()
            if string != u'':
                for field in string.split(','):
                    self.Fields.append(field.strip())
        else:
            self.FieldsCount = len(self.Fields)
        #Получаем элементы дата сэта
        dogflag = False
        ItemStringsList = []
        for string in ListOfStrings[index+1:]:
            if (string.strip() != u'' and string.strip()[:8] != u'!=======') or dogflag:
                ItemStringsList.append(string)
                if dogflag and (u'"@,' in string or string.strip()[-2:] == u'"@'):
                    dogflag = False
                if string.strip()[-2:] != u',\\' and string.strip()[-2:] == u'@"':
                    dogflag = True
                if string.strip()[-2:] != u',\\' and not dogflag:
                    Attributes = self.GetAttributesFromList(ItemStringsList)
                    self.Items.append(ItemOfDataSet(Attributes))
                    ItemStringsList = []
        else:
            self.ItemsCount = len(self.Items)

    def GetAttributesFromList(self, List):
        Attributes = {}
        Data = []
        dogflag = False
        dogstring = u''
        for string in List:
            if dogflag and (u'"@,' in string or string.strip()[-2:] == u'"@'):
                dogflag = False
                cutplace = string.find('"@')
                dogstring = dogstring + string[:cutplace]
                string = string[cutplace + 3:]
                Data.append(dogstring)
                dogstring = u''
                if string.strip() == u'\\' or string.strip() == u'':
                    continue
            elif dogflag:
                dogstring = dogstring + string + u'\n'
            elif string.strip()[-2:] != u',\\' and string.strip()[-2:] == u'@"':
                dogflag = True
            if not dogflag:
                ds = self.SplitString(string)
                if u'\\' in ds:
                    ds.remove(u'\\')
                if len(ds) > 1 or (len(ds) == 1 and ds[0] <> u''):
                    for d in ds:
                        if d[0] == u'"' and d[-1] == u'"':
                            d = d[1:-1]
                        elif d[:2] == u'@"' and d[-2:] == u'"@':
                            d = d[2:-2]
                        Data.append(d)
        if len(Data) != self.FieldsCount:
            print u'Проблема в датасете ', self.Name
            print Data
        else:
            for i in range(self.FieldsCount):
                Attributes[self.Fields[i]] = Data[i]
        return Attributes

    def SplitString(self, st):
        open = False
        close = False
        items = []
        item = ''
        i = 0
        for c in st:
            if c == '"' and close:
                close = False
            elif c == ',' and (not open or close):
                items.append(item)
                item = ''
                open = False
                close = False
                i += 1
                continue
            elif not open and c == '"':
                open = True
            elif open and c == '"':
                close = True
            item = item + c
            i += 1
            if len(st) == i:
                items.append(item)
        return items

    def MakeXlsTable(self, fields = []):
        table = []
        table.append(self.Fields)
        textfiles = {}
        textfileindex = 1
        for item in self.Items:
            itemtableline = []
            for field in self.Fields:
                textfile = []
                if u'\n' in item.__dict__[field]:
                    filename = self.Name[1:] + '_' + str(textfileindex)
                    itemtableline.append('File_@"@_'+ filename)
                    textfileindex += 1
                    for line in item.__dict__[field].split(u'\n'):
                        line = line + u'\n'
                        textfile.append(line.encode('cp1251'))
                    textfiles[filename] = textfile
                else:
                    itemtableline.append(item.__dict__[field])
            table.append(itemtableline)
        return ({u'1|' + self.Name: table}, textfiles)


class ItemOfDataSet(object):
    u'''Элемент датасэта'''
    def __init__(self, Attributes):
        for attr, value in Attributes.items():
            self.__dict__[attr] = value

    def SayWhatYouHave(self):
        for attr, value in self.__dict__.items():
            print attr, value
        print u'*******************************************'

if __name__ == "__main__":
    b = XLSbase('C:\\Share\\XLS')
    #b.SayWhatYouHave()