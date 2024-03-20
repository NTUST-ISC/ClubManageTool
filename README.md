# ClubManageTool
Automation tool for adding member to club manage system.

# Installation
## Download
- [Google Chrome](https://www.google.com/intl/zh-TW/chrome/)  
- [Chrome Driver](https://chromedriver.chromium.org/downloads)

chrome driver `.zip` -> unzip and put it into this directory

## Python
- [Python](https://www.python.org/downloads/)
- install dependencies

```sh
pip install -r requirements.txt
```

# Usage
## Account
1. create `.env` and fill in `USERNAME` and `PASSWORD`  
2. fill in the `.xlsx` file  
 (if a memeber is `社員`, its `職位` column can be empty)
3. run the script  

```sh
python main.py
```

