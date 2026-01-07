# Excel Logistics Automation


The VBA modules in this repository were made for a local delivery company to automate the fetching of data from Excel files.
FileSelectionMac is a modified version of the code written by Ron de Bruin (https://macexcel.com/examples/filesandfolders/selectfiles/).
GetPrealertAndLentoRekka is written by me.
The Excel files contain delivery data and new files arrive every day via email.


The modules I wrote/modified work together with an email extension that automatically downloads those Excel files into their respective folders.
Then, with the click of a button, the data is fetched from those files and from X previous workdays as well (configured by the user). 
Finally, the data is pasted into the Excel file that is used to process the data.


## Features

- Automatic/Manual selection of files
- Configurable behaviour via a config worksheet
- Cross platform support (macOS/Windows)
- User friendly error handling
- Performance optimization
- Workday aware data fetching
- Dynamic file imports
