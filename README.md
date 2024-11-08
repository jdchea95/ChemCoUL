# ChemCoUL
This repository contains the output files from the Chemical Condition of Use Locator (ChemCoUL) algorithm. 

Each folder contains two qualitative mapping .PNG files that illustrates the chemical flow mapping through commerce. There is also a pdf generated, containing a facility summary report of all of the chemical-related data and facilities registered within the Facility Registry Service (FRS). 

The Condition-of-use-TRI1b-CDR-PUCS-NAICS-CASRN.xlsx file contains the raw data acquired from the ChemCOUL algorithm (including database search + crosswalked information)

The Condition-of-use file was created from two initial search operations that yield (1) TRI_CDR_REVERSED_CASRN.xlsx, containing the raw data from the Toxics Release Inventory and Chemical Data Reporting databases and (2) the Product_Use_Information_NAICS_Crosswalked_CASRN.xlsx file, containing the raw data acquired from the Chemical and Products Database (CPDat), crosswalked with manufacturing industry NAICS codes. 

The qualitative_chemical_flow_mapping_summary_CASRN.xlsx is final spreadsheet output after all of the data have been processed. All FRS-registered facilities have been assigned a unique code (e.g., FRS1(I), FRS2(I)) to maintain confidentiality, then placed in different sheets for creating qualitative mapping flow diagram (.PNG files).
