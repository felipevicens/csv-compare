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

import pandas as pd
import os, shutil

def convert_excel_csv(filename):
    '''
    This function convert a xls or xlsx file into a multiple csv files in temp directory
    Return a list of csv files in the tmp directory
    '''
    excel = pd.ExcelFile(filename)
    sheets = excel.sheet_names
    files_location = f"/tmp/{filename}"
    os.mkdir(files_location)
    for sheet in sheets:
        sheet_content = pd.read_excel(excel, sheet)
        sheet_filename = f'/tmp/{filename}/{sheet}.csv'
        sheet_content.to_csv(sheet_filename, index = False, header = True)
    return files_location

def clean_temp(directories):
    for directory in directories:
        try:
            shutil.rmtree(directory)
        except Exception as e:
            print(f"ERROR: {e}")