import traceback
import argparse
from sheetprotect import *
from constans import *


print("""
#####  #####  #####   ####   ####  ######
##  ## ##     ##  ## ##  ## ##  ##   ##
#####  ####   #####  ##  ## ##  ##   ##
##  ## ##     ##  ## ##  ## ##  ##   ##
##  ## #####  #####   ####   ####    ##\n""")

print("[#] Remove Sheet Protection v. 1.9\n")
parser = argparse.ArgumentParser()
parser.add_argument("--file","-f",type=str, help="file to edit")
parser.add_argument("--info","-i", action="store_true",
                                    help="print info")   
parser.add_argument("--nohide","-n", action="store_true",
                                    help="remove type \"hidden\" from sheets")   
parser.add_argument("--nopatch","-p", action="store_true",
                                    help="cancel removed protection")   
parser.add_argument("--struct","-s", action="store_true",
                                    help="remove structure")  
args = parser.parse_args()
file_name = args.file

try:
        print("[+] Extracting files...")
        extract(file_name)
        
        if args.info:
                get_infomation = get_info(PATH_WORKBOOK)
                lenght = len(get_infomation)
                print("\n[+] All sheets:", lenght)
                hidden_count = 0
                for i in range(lenght):
                        info = get_infomation[i]
                        name = info["name"]
                        state = info["state"]
                        if state == "hidden": 
                                hidden_count += 1
                        is_protect = info["is_protect"]
                        print("\t[>] {}. {}, state: {}, protected: {}".format(i + 1, 
                                                                        repr(name), 
                                                                        state,
                                                                        is_protect))
                print("\n\t[>] Hidden sheets:", hidden_count)
                structure_lock = find_tag(PATH_WORKBOOK, tag=TAG_PROTECTION_WORKBOOK)
                structure_lock = "Yes" if structure_lock else "No"
                print("[>] Structure protection:", structure_lock)
        else:
                if not args.nopatch:
                        print("[+] Patching sheets...")
                        patching()
                if args.nohide:
                        print("[+] Removing type \"hidden\" sheets...")
                        if remove_hide():
                                print("\t[>] Successfully")
                        else:
                                print("\t[>] Unsuccessfully")
                if args.struct:
                        print("[+] Removing structure protection...")
                        remove = remove_tag(PATH_WORKBOOK, tag=TAG_PROTECTION_WORKBOOK)
                        if remove and remove != "Error":
                                print("\t[>] Successfully")
                        else:
                                print("\t[>] Unsuccessfully")

                print("[+] Compressing patched file")
                compress(file_name)
        print("[+] Removing extract...")
        remove_extract(file_name)
        print("\n[*] Done!")

except Exception as ex:
    print("[-] Error")
    traceback.print_exc()
    remove_extract(file_name)