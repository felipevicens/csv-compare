#!/usr/bin/python3
#  Copyright (c) 2020 Felipe Vicens
# ALL RIGHTS RESERVED.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
Excel Workbook row-by-row comparison.

Two options are available. Process by comparing all sheets between two
xls/xlsx files or comparing all csv files between 2 folders.

When --process is folder. It receive 2 arguments:
- First: The folder containing the old csv files to be compared
- Second: The folder containing the new csv files to be compared

When --process is file. It receive 2 arguments:
- First: The file path with the old xls/xlsx version
- Second: The file path with the new xls/xlsx version

[OPTION] --ui Additionally you can make use of tkdiff for visual comparison
[OPTION] --clean To skip the context around the change

"""
import hashlib, glob, os, sys
from os import listdir
from os.path import isfile, join
import difflib
import click
import subprocess
from .excel import convert_excel_csv, clean_temp

# Importing color
try:
    from colorama import Fore, Back, Style, init
    init()
except ImportError:  # fallback so that the imported classes always exist
    class ColorFallback():
        __getattr__ = lambda self, name: ''
    Fore = Back = Style = ColorFallback()

@click.command()
@click.option('--process', default='file', help='By "file" or "folder"')
@click.option('--ui', default=False, is_flag=True, help="[OPTIONAL] Use tkdiff for visual comparison. Requires tkdiff installation in the OS")
@click.option('--clean', default=False, is_flag=True, help="[OPTIONAL] Skip the context around the change")
@click.argument('old')
@click.argument('new')
def cli(process, old, new, ui, clean):
    """
        Excel Workbook row-by-row comparison.
        
        \b
        Process by comparing all sheets between two xls/xlsx files or 
        comparing all csv files between 2 folders.

        \b
        When option --process is "folder". It receive 2 arguments:
        - First: The folder containing the old csv files to be compared
        - Second: The folder containing the new csv files to be compared

        \b
        When option --process is "file". It receive 2 arguments:
        - First: The file path with the old xls/xlsx version
        - Second: The file path with the new xls/xlsx version
    """

    missing_sheets = []
    different_sheets = []

    def color_diff(diff, clean):
        """
            Color the differences from difflib library
            :param fname: file path
            :return: checksum string
        """
        for line in diff:
            if line.startswith('+'):
                yield f"{Fore.GREEN}{line}{Fore.RESET}"
            elif line.startswith('-'):
                yield f"{Fore.RED}{line}{Fore.RESET}"
            elif line.startswith('^'):
                yield f"{Fore.BLUE}{line}{Fore.RESET}"
            elif line.startswith('@'):
                yield f"{Fore.BLUE}{line}{Fore.RESET}"
            else:
                if not clean:
                    yield line

    def md5(fname):
        """
            Checksum generator
            :param fname: file path
            :return: checksum string
        """
        hash_md5 = hashlib.md5()
        with open(fname, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def folder_process(old, new, ui, clean):
        missing_sheets = []
        files_old = [f for f in listdir(old) if isfile(join(old, f))]
        files_new = [f for f in listdir(new) if isfile(join(new, f))]

        if set(files_old).difference(set(files_new)) != set():
            missing_sheets = list(set(files_old).difference(set(files_new)))

        # Main
        # Checking the files located in old folder against files in new_folder

        for file_folder in files_old:
            if file_folder in files_new:
                if md5(f"{old}/{file_folder}") == md5(f"{new}/{file_folder}"):
                    print(f"{Fore.MAGENTA}[NO_CHANGES]:{Fore.RESET} {file_folder}")
                else:
                    print(f"{Fore.YELLOW}[DIFFERENT]:{Fore.RESET} {file_folder}")
                    with open(f"{old}/{file_folder}", encoding="UTF-8", errors="ignore") as f_old:
                        f_old_text = f_old.readlines()
                    with open(f"{new}/{file_folder}", encoding="UTF-8", errors="ignore") as f_new:
                        f_new_text = f_new.readlines()
                    
                    # Calculate the differences
                    if ui == True:
                        try:
                            subprocess.run(["/usr/bin/tkdiff", f'{old}/{file_folder}', f'{new}/{file_folder}'])
                        except Exception as e:
                            print(f"\n{Fore.RED}ERROR: {e}\n{Fore.YELLOW}This option is only valid for linux and requires tkdiff. Please install tkdiff by:{Fore.RESET}\n\nsudo apt-get install tkdiff\n")
                            exit(1)

                    else:
                        diff = difflib.unified_diff(f_old_text, f_new_text, fromfile=f"{old}/{file_folder}", tofile=f"{new}/{file_folder}", lineterm='')
                        diff = color_diff(diff, clean)
                        print('\n'.join(diff))

                    # Add different files into a list
                    different_sheets.append(file_folder)
        return missing_sheets, different_sheets

    def file_process(filename):
        print(f"Procesing File: {filename}")
        return convert_excel_csv(filename)
    
    if process == 'folder':
        try:
            missing_sheets, different_sheets = folder_process(old, new, ui, clean)
        except Exception as e:
            print(f"ERROR: {e}")
            exit(1)

    elif process == 'file':
        try:
            csv_old = file_process(old)
            csv_new = file_process(new)
            missing_sheets, different_sheets = folder_process(csv_old, csv_new, ui, clean)
        except Exception as e:
            print(f"ERROR: {e}")
            clean_temp([f"/tmp/{old}", f"/tmp/{new}"])
            exit(1)
        finally:
            clean_temp([f"/tmp/{old}", f"/tmp/{new}"])
    else:
        print(f"option -p or --process must be file or folder. Used: {process}")


    # Report
    print(f"\n\n{Fore.YELLOW}*** REPORT ***{Fore.RESET}")
    print(f"Different Sheets: {different_sheets}")
    print(f"Missing Sheets: {missing_sheets}")