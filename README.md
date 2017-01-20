#Environmental-Data-Formatting 
A subset of the tools I use for formatting environmental analytical data (soil, groundwater and air) written in VBA for Excel.

This repository has one excel file and three VBA modules:

The excel file is Example.xlsm, an example of the data format delivered from the lab. This excel file has the standard template and macros to be able to run "AAConvertLabDataToReportFormat", a macro used to (as the name implies) convert lab data into a standard report format

The three vba modules are functions.bas, individual_screening_subs.bas and convert_labdata_to_report_format.bas . convert_labdata_to_report_format uses the classes and functions from individual_screening_subs which uses the functions from... you guessed it, functions.
