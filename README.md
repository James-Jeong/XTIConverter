# 🍓 XLSX To INI Converter
#
### 1. Env: Python3
#
### 2. Module
  #### 1) openpyxl
    1] sudo pip3 install openpyxl
    2] source compile
      1} https://pypi.org/project/openpyxl/#files 에서 openpyxl-x.x.x.tar.gz 파일 다운로드 (File type: Source)
      2} cd openpyxl-x.x.x
      3} python3 setup.py install
---
  #### 2) configparser
    - sudo pip3 install configparser
#
### 3. How to use
  #### 1) Common argument
    - argv[0]: XTIConverter.py
---
  #### 2) Config mode
    1] Argument
      - argv[1]: {config path}
    2] command: python3 argv[0] argv[1]
    3] example: python3 XTIConverter.py $HOME/XTIConverter/config/default_conf.ini
---
  #### 2) Parameter mode
    1] Arguments
      1} argv[1]: {xlsx path}
      2} argv[2]: {xlsx sheet name}
      3} argv[3]: {ini path}
    2] command: python3 argv[0] argv[1] argv[2] argv[3]
    3] example: python3 XTIConverter.py $HOME/XTIConverter/resources/test1.xlsx Sheet1 $HOME/XTIConverter/resources/network_conf.ini
