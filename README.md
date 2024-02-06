# Vpdive CSV/Excel to CSV Transformer

Simple python scripts that allows to transform/extract from the VPDive Member list (Excel sheet) into CSV files.
This allows to e.g. update the CERN egroup lists.

It consists of 2 tools:
1. Egroup extractor (members or guests)
2. Medical certificate (members)

## Vpdive CSV/Excel to Egroup CSV Transformer

How to:

1 Get Excel sheet
  - Go on VPDive -> Admin section -> Profile des members
  - Download there the excel sheet (telecharger)

2 **(This step is not needed anymore!)** Transform Excel sheet to CSV 
- Open the excel sheet and save it as csv (e.g. current_members.csv)
    (In Microsoft Excel this can be done with save as/export)


3 Run the script
- `python3 ./createCSVforEgroups.py -i current_members.csv`
  
OR
- `python3 ./createCSVforEgroups.py -i current_members.csv -m guests`
  (egroup list for non-members e.g. for csc-club-friends)

The script will create a new csv named `egroup_<in-name>.csv`
(e.g. here `egroup_current_member.csv` or `guests_current_member.csv`)
that fulfills the data format requirements for the egroup import

4 Update the egroup (need admin rights):
- Go to e-groups.cern.ch and search for club-csc-members 
- Go to Members -> Import
- Select the following
  - Strategy: `Overwrite`
  - Filename:  `egroup_<in-name>.csv`
- Press "upload file"


## Vpdive CSV/Excel to CACI CSV Transformer

Output is a CSV file with format (contains no header):

`last name, first name, medical certificate expire date`

Run: `python3 ./createCSVforEgroups.py -i current_members.csv -m mode`

Outname: `caci_<in-file>.csv`


## Needed Python packages

```
xlrd
pandas
numpy
argcomplete
```

Note: To enable auto complete functionality for file discovery, please
      have a look here: https://pypi.org/project/argcomplete/#activating-global-completion