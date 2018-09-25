#!/usr/bin/env python
# -*- coding: utf-8 -*-
from MyProject import Project

#Путь к файлу экспорта из FastTools. Одинарный слеш "\", заменяется на двойной "\"
XLSfolerPath = u'C:\\Share\\XLS'
#Путь к папке прокта. Одинарный слеш "\", заменяется на двойной "\"
ResultPath = u'C:\\Share\\TestImport'

Result = Project(XLSfolerPath, ResultPath)
Result.main()

print (u'Первый этап выполнен, внесите необходимые изменения')