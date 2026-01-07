# 📊 Excel Logistics Automation


This repository contains VBA modules created for a local delivery company to automate the collection and processing of delivery data from Excel files. 

The `FileSelectionMac` module is a modified version of Ron de Bruin’s code (https://macexcel.com/examples/filesandfolders/selectfiles/). The `GetPrealertAndLentoRekka` module was written by me.

The system works with an email extension that automatically downloads new Excel files into designated folders. With a single click, the modules fetch data from the latest files and a configurable number of previous workdays, then paste it into a central Excel workbook for processing.



## ✨ Features

- Automatic/Manual selection of files
- Configurable behaviour via a config worksheet
- Cross platform support (macOS/Windows)
- User friendly error handling
- Performance optimization
- Workday aware data fetching
- Dynamic file imports from arbitrary number of files and dates
- Compatible with large datasets (used to fetch 50 000 rows)
