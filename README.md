# Subset of tools for formatting analytical data

A subset of the tools I use for formatting environmental analytical data (soil, groundwater and air) written in VBA for Excel.

This repository has one excel file and three VBA modules:

The excel file is `Example.xlsm`, an example of the data format delivered from the lab.

The three vba modules are:
 - `functions.bas`: Helper functions.
 - `individual_screening_subs.bas`: Functions for screening data against regulatory standards to understand what samples/compounds exceed.
 - `convert_labdata_to_report_format.bas`: The tool to convert the lab data to report ready format (uses functions from the two modules above).
