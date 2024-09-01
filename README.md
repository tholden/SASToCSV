# SASToCSV

Code for converting SAS files (including SAS V6 SD2 files) to CSV.

This only runs on Windows.

If you do not have SAS installed, then (assuming you want to process SD2 files) you first need to install the 32 bit version of the August 2014 edition of Release 9.4 of the SAS Providers for OLE DB from https://support.sas.com/downloads/package.htm?pid=1265

(If you only want to process newer SAS files, then you can also use more recent 64 bit versions of the SAS Providers for OLE DB.)

The approach is derived from the one described here: https://blogs.sas.com/content/sasdummy/2012/04/12/build-your-own-sas-data-set-viewer-using-powershell/

The code is derived from the PowerShell script provided in this repository: https://github.com/sassoftware/sas-inttech-samples
