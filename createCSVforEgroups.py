import pandas as pd
import numpy as np
import argparse
import argcomplete
import glob
import textwrap
import xlrd #for excel

# For description and usage please look in function "parse_arguments()"

def getFiles(prefix, parsed_args, **kwargs):
  files = glob.glob(prefix + "*")
  return files


# TODO(Laura) check if that works - definitely not working on windows
allowed_modes = ["egroup", "caci", "guests"]
def getModes(prefix, parsed_args, **kwargs):
  return [ele for ele in allowed_modes if ele.startsWith(prefix) == True]

def parse_arguments():
  parser = argparse.ArgumentParser(
    prog='python3 ./createCSVforEgroups.p -i <vpdive_memberlist>.csv -m <mode>',
    formatter_class=argparse.RawDescriptionHelpFormatter,
    description=textwrap.dedent("""
Vpdive CSV/Excel to Egroup CSV Transformer
==========================================
How to:
1.a Go on VPDive -> Admin section -> Profil des members
1.b Download there the excel sheet (telecharger)

Step 2 is not needed anymore. But in case the xls cannot be read:
2. Open the excel sheet and save it as csv (e.g. current_members.csv)
    (In Microsoft Excel this can be done with save as/export)

3a. Run this script
    python3 ./createCSVforEgroups.py -i current_members.csv
OR
3b. Run this script
    python3 ./createCSVforEgroups.py -i current_members.csv -m guests
    (egroup list for non-members e.g. for csc-club-friends)

 The script will create a new csv named egroup_<in-name>.csv
 (e.g. here egroup_current_member.csv or guests_current_member.csv)
 that fulfills the data format requirements for the egroup import

4. Update the egroup (need admin rights):
 a. Go to e-groups.cern.ch and search for club-csc-members 
 b. Go to Members -> Import
 c. Select the following
   Strategy: Overwrite
   Filename:  egroup_<in-name>.csv
 d. Press "upload file"


Vpdive CSV/Excel to CACI CSV Transformer
==========================================
To get a list of:
last name, first name, medical certificate expire date

use:
python3 ./createCSVforEgroups.py -i current_members.csv -m mode

outname:
caci_<in-file>.csv

Note: To enable auto complete functionality for file discovery, please
      have a look here: https://pypi.org/project/argcomplete/#activating-global-completion
    """),
    epilog="Author: HereThereBeDragons (2023-2024)")
  parser.add_argument('-i', '--input',
                      help='VPdive Member List as CSV file',
                      required=True).completer = getFiles
  parser.add_argument('-m', '--mode',
                      choices = allowed_modes,
                      help='egroup (default, for members), caci (for members) or guests (for egroup list of guests)',
                      default = "egroup",
                      required=False).completer = getModes
  argcomplete.autocomplete(parser)
  return parser.parse_args()

def getEgroup(df): 
  arr_egroups = []

  for ele in df["Email"]:
      arr_egroups.append(["E","",ele])
  
  return arr_egroups

def getCaci(df): 
  arr_caci = []

  df1 = df[["Nom", "Prénom", "CACI"]]

  for ele in df1.values:
    arr_caci.append(ele.flatten())
  
  return arr_caci

def selectGroup(df, are_members):
  # find idx of "Liste des invités au site" (the members will be before it)
  for idx, ele in enumerate(df.iloc[:,0].astype(str)):
    if "au site" in ele:
        print("Found end of member section at column:", idx, "with value '" + ele + "'")
        delete_idx = idx # we have 2 nan/empty rows before and 1 nan/empty row behind

  if are_members == True:
    selected_df = df.iloc[:(delete_idx-2)]  
  else:
    selected_df = df.iloc[delete_idx+2:]
  
  return selected_df



if __name__ ==  "__main__":
    parsed_args = parse_arguments()
    input_file = parsed_args.input #"V:/Diving/2023/Secretary/liste_des_membres__26_09_2023.csv"
    mode = parsed_args.mode

    if len(input_file.split("/")) == 1:
        input_file = "./" + input_file

    if input_file.split(".")[-1] == "csv":
      df = pd.read_csv(input_file, encoding="iso-8859-1", skiprows=3)
    elif input_file.split(".")[-1] == "xls" or input_file.split(".")[-1] == "xlsx":
      workbook = xlrd.open_workbook(input_file, ignore_workbook_corruption=True)
      df = pd.read_excel(workbook, sheet_name = 0, skiprows=3)
    else:
      print("No valid input file format. Support: .csv, .xls and .xlsx")
      exit(0)
    
    are_members = True

    if mode == "guests":
      are_members = False

    df = selectGroup(df, are_members)

    if "egroup" in mode or "guests" in mode:
      final_selection = getEgroup(df)
    elif "caci" in mode:
      final_selection = getCaci(df)
    
    outname = input_file.rsplit("/", 1)[0] + "/" + mode + "_" + input_file.rsplit("/", 1)[1].rsplit(".", 1)[0] + ".csv"

    print("Create " + mode + " csv: ", outname)
    pd.DataFrame(final_selection).to_csv(outname, header=False, index=False)
