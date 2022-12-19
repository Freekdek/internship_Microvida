# internship_Microvida
## Introduction
Python scripts for internship Freek de Kreek

Containts qc_collector_v1.py; a Python script that collects QC values from trimreports produced by CLC and parses in a QC template.
seqsphere_report_rename_v1.py; a Python script that renames files with a MWGS_id to GLIMS_id + "1" + isolatenr based on a metadata file.


## How to use
#### qc_collector_v1.py:
Reads all Excel files (qc trim reports) of specified  directory (Input) and collects all the required values (in Output):
- Any nucleotide, Count, N50, Maximum, b Reads, Contigs, Matched
- Every row is read, checks for the afore mentioned values and when found writes the values in the corresponding collumn
- Duplicates are not allowed

```
py qc_collector_v1.py -i input -o output

-i  --Input : provide path to trim reports
-o  --Output  : provide path to the QC template
```

#### seqsphere_report_rename_v1.py:
For each file:
- Convert file name to MWGS number only and put it in dictionary
- MWGS number found -->
- Save GLIMS ID by adding to dictionary value (list)
- Store isolate number by adding to dictionary value (list)
- Write file name --> GLIMS ID + “1” + isolate number + “wgs.pdf”

```
py seqsphere_report_rename_v1.py -i input -d metadata -o output

-i  --Input : provide path to the files that need to be renamed
-d  --Database  : provide path to metadata file
-o  --Output  : provide path where to copy the renamed files to
-v  --Verbose : specifying "True" for this argument will turn on debugging
```

## Dependencies
**Python Packages:**
- os
- openpyxl
- argparse
- timeit
- shutil
