# Data Management System

This is a business solution with built-in automation for KPI estimations and data handling in MS Excel (data reporting) and a relational database service provider (data storage in SQL server or MS Access).

## Overview

Transactional data is collected in 45 different worksheets in Excel and Access and later It has to be centralized and consolidated in a single data base repository. This task is automated through VB Macros and ADO connections executed in windows forms apps, which continously populate workbook datasets validating data consistency and cleaning. Additionally, the userform tracks data that has been overwritten, deleted or added by any user providing a dynamic log of changes.

## Packages
The following package classes must be imported in the VBA developer tab to execute the data repository userforms and macros:

1. Microsoft Excel 16.0 Object Library
2. VBA OLE Automation library
3. Microsoft Windows Forms 2.0 Object library
4. AppxManagerLib
5. Microsoft ActiveX Data Objects 6.1 library



