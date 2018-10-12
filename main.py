#!/usr/bin/env python
# -*- coding: utf-8 -*-
from MyProject import Project

#Путь к файлу экспорта из FastTools. Одинарный слеш "\", заменяется на двойной "\"
XLSfolerPath = u'C:\\Temp\\KOS_XLS_NEW'
#Путь к папке прокта. Одинарный слеш "\", заменяется на двойной "\"
ResultPath = u'C:\\Temp\\KOS_XLS_NEW\\REUSULT'

Result = Project(XLSfolerPath, ResultPath)
Result.main()

print (u'All done')