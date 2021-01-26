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
  ### 2) Argument
    - argv[0]: XTIConverter.py or ITXConverter.py
    - argv[1]: {config path}
  ### 3) Command
    - python3 argv[0] argv[1]
  ### 4) Example
    - python3 XTIConverter.py $HOME/XTIConverter/config/default_conf.ini
