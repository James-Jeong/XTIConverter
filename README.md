# 🍓 XLSX & INI Converter
### 1) XLSX To INI
### 2) INI To XLSX
#
## 1. Env: Python3
#
## 2. Module
  ### 1) openpyxl
    1] sudo pip3 install openpyxl
    2] source compile
      1} https://pypi.org/project/openpyxl/#files 에서 openpyxl-x.x.x.tar.gz 파일 다운로드 (File type: Source)
      2} cd openpyxl-x.x.x
      3} python3 setup.py install
---
  ### 2) configparser
    - sudo pip3 install configparser
#
## 3. How to use
  ### 1) Installation path: $HOME directory
  ### 2) Common argument
    - argv[0]: XTIConverter.py or ITXConverter.py
---
  ### 3) Config mode
    1] Argument
      - argv[1]: {config path}
    2] Command: python3 argv[0] argv[1]
    3] Example: python3 XTIConverter.py $HOME/XTIConverter/config/default_conf.ini
---
  ### 4) Parameter mode
    1] XTI arguments
      1} argv[1]: {xlsx path}
      2} argv[2]: {xlsx sheet name}
      3} argv[3]: {ini path}
    2] ITX arguments
      1} argv[1]: {ini path}
      2} argv[2]: {xlsx path}
      3} argv[3]: {xlsx sheet name}
    3] Command: python3 argv[0] argv[1] argv[2] argv[3]
    4] Example1: python3 XTIConverter.py $HOME/XTIConverter/resources/test1.xlsx Sheet1 $HOME/XTIConverter/resources/network_conf.ini
    5] Example2: python3 ITXConverter.py $HOME/XTIConverter/resources/network_conf.ini $HOME/XTIConverter/resources/test1.xlsx Sheet1
