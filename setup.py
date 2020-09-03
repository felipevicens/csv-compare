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

from setuptools import setup, find_packages

_description = 'Tool for comparing csv files between 2 folders'

setup(name='csv_compare',
      license='Apache License, Version 2.0',
      version='0.1',
      url='https://github.com/felipevicens/csv-compare',
      author='Felipe Vicens',
      author_email='fjvicens@edgecloudlabs.com',
      description=_description,
      include_package_data=True,
      packages=find_packages(),
      install_requires=['Click', 'colorama'],
      entry_points="""
          [console_scripts]
          csv-compare=csv_compare.scripts.compare:cli
          """,
      )