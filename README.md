# PockEthernet-Report-Scanner
Simple powershell script to parse through PockEthernet CDP Switch Port Scans done in the field by technicians.  

## Dependancies 
Module - PSWritePDF
The script _should_ install it if not present, but if not, try "Install-Module -Name PSWritePDF -Force"

PockEthernet is a "A smartphone connected Ethernet network analyzer & cable tester that fits into your pocket".  PockEthernet device creates reports on network drops such as wiremapping, CDP, LLDP, and other functions. https://pockethernet.com/

## Function
This script parses through reports that are typically generically named such as "Pockethernet report 2021-06-14 - 14-12-56" which does not give much information into the report without opening it.

The script will filter through reports found in C:\Users\USERID\scans folder and will extract information.  After information has been parsed, the script will rename that report file according to the information found in the report.

This allows the user to parse through hundreds of network port scans very quickly and rename them to a naming schema that reflects the information contained.

The script will also export all information parsed into a CSV file as well as a long term CSV LOG file in the users C:\Users\USERID.

## Currently the script extracts the following information
![image](https://user-images.githubusercontent.com/40606082/132001702-2d102ebf-fb79-4bd3-bc97-682a9e4f7f24.png)


## Duplicate Detection

The script will detect reports with duplicate information (exact copy) and report the total number of indexed reports and duplicates.  Duplicates will be named with the tag (DUPLICATE)







