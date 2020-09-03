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
Excel CSV line-by-line comparison.

This is script is part of the excel sheet comparison. Before

It receive 2 arguments:
- First: The folder containing the old files to be compared
- Second: The folder containing the new files to be compared

"""
import hashlib, glob, os, sys
from os import listdir
from os.path import isfile, join
import difflib
import click

# Importing color
try:
    from colorama import Fore, Back, Style, init
    init()
except ImportError:  # fallback so that the imported classes always exist
    class ColorFallback():
        __getattr__ = lambda self, name: ''
    Fore = Back = Style = ColorFallback()

@click.command()
@click.argument('folder_old')
@click.argument('folder_new')
def cli(folder_old, folder_new):
    """
    \b
    Steps in Microsoft Excel:

    \b
    1) Press Alt + F11
    2) Right clic on a sheet and select Insert > module
    3) Paste the code below:

    \b
    Sub SplitWorkbook()
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim xWs As Worksheet
    Dim xWb As Workbook
    Dim FolderName As String
    Application.ScreenUpdating = False
    Set xWb = Application.ThisWorkbook
    DateString = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    FolderName = xWb.Path & "\" & xWb.Name & " " & DateString
    MkDir FolderName
    For Each xWs In xWb.Worksheets
        xWs.Copy
        FileExtStr = ".csv": FileFormatNum = 6
        xFile = FolderName & "\" & Application.ActiveWorkbook.Sheets(1).Name & FileExtStr
        Application.ActiveWorkbook.SaveAs xFile, FileFormat:=FileFormatNum
        Application.ActiveWorkbook.Close False
    Next
    MsgBox "You can find the files in " & FolderName
    Application.ScreenUpdating = True
    End Sub
    \b
    4) Press F5 to execute the code. It Takes some time. You can grab a coffe :)
    5) The files will be generated in the Spreadsheet saved folder.

    """

    files_folder_old = [f for f in listdir(folder_old) if isfile(join(folder_old, f))]
    files_folder_new = [f for f in listdir(folder_new) if isfile(join(folder_new, f))]

    if set(files_folder_old).difference(set(files_folder_new)) != set():
        missing_files = list(set(files_folder_old).difference(set(files_folder_new)))
    else:
        missing_files = None

    files_differents = []

    def color_diff(diff):
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

    # Main
    # Checking the files located in old folder against files in new_folder

    for file_folder in files_folder_old:
        if file_folder in files_folder_new:
            if md5(f"{folder_old}/{file_folder}") == md5(f"{folder_new}/{file_folder}"):
                print(f"{Fore.MAGENTA}File [SAME]:{Fore.RESET} {file_folder}")
            else:
                print(f"{Fore.YELLOW}File [DIFF]:{Fore.RESET} {file_folder}")
                with open(f"{folder_old}/{file_folder}", encoding="UTF-8", errors="ignore") as f_old:
                    f_old_text = f_old.readlines()
                with open(f"{folder_new}/{file_folder}", encoding="UTF-8", errors="ignore") as f_new:
                    f_new_text = f_new.readlines()
                
                # Calculate the differences
                diff = difflib.unified_diff(f_old_text, f_new_text, fromfile=f"{folder_old}/{file_folder}", tofile=f"{folder_new}/{file_folder}", lineterm='')
                diff = color_diff(diff)
                print('\n'.join(diff))
                
                # Add different files into a list
                files_differents.append(file_folder)

    # Report
    print("\n\n*** REPORT ***")
    print(f"Files differents: {files_differents}")
    print(f"Missing files: {missing_files}")