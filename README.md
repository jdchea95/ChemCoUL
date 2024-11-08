# ChemCoUL
This repository contains the output files from the Chemical Condition of Use Locator (ChemCoUL) algorithm. The Python code used to run the algorithm is available in the Supporting Information, Section B. Due to the file size restriction in GitHub, dependency files that can be obtained from publicly available database were not uploaded. However, please reach out to the corresponding author (Chea.John@epa.gov) if you require assistance with acquiring the dependency files.

Data Availability Statement
The data that support the findings of this study are openly available in the Chemical Data Reporting database at https://www.epa.gov/chemical-data-reporting/access-chemical-data-reporting-data, Toxics Release Inventory at https://www.epa.gov/toxics-release-inventory-tri-program/tri-basic-plus-data-files-calendar-years-1987-present, and Chemical and Products Database at https://www.epa.gov/chemical-research/chemical-and-products-database-cpdat. 


Regarding the result of each chemical (e.g., Ammonia),
Each chemical folder contains two qualitative mapping .PNG files that illustrates the chemical flow mapping through commerce. There is also a pdf generated, containing a facility summary report of all of the chemical-related data and facilities registered within the Facility Registry Service (FRS). 

The Condition-of-use-TRI1b-CDR-PUCS-NAICS-CASRN.xlsx file contains the raw data acquired from the ChemCOUL algorithm (including database search + crosswalked information)

The Condition-of-use file was created from two initial search operations that yield (1) TRI_CDR_REVERSED_CASRN.xlsx, containing the raw data from the Toxics Release Inventory and Chemical Data Reporting databases and (2) the Product_Use_Information_NAICS_Crosswalked_CASRN.xlsx file, containing the raw data acquired from the Chemical and Products Database (CPDat), crosswalked with manufacturing industry NAICS codes. 

The qualitative_chemical_flow_mapping_summary_CASRN.xlsx is final spreadsheet output after all of the data have been processed. All FRS-registered facilities have been assigned a unique code (e.g., FRS1(I), FRS2(I)) to maintain confidentiality, then placed in different sheets for creating qualitative mapping flow diagram (.PNG files).

