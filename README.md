# ClubManageTool
Automation tool for adding member to club manage system.

# Installation
## Pre-Download
- [Google Chrome](https://www.google.com/intl/zh-TW/chrome/)
- [Chrome Driver](https://chromedriver.chromium.org/downloads)

Unzip chrome driver's `.zip` and put `chromedriver` into this directory.  
![image](https://github.com/NTUST-ISC/ClubManageTool/assets/55608737/3247e622-1754-4396-8c9f-e23795580dc4)


## Python
- Install [Python](https://www.python.org/downloads/).
- Install dependencies.

```sh
pip install -r requirements.txt
```

# Usage
## Account
### .env
Create `.env` and fill in `USERNAME` and `PASSWORD`.  
![image](https://github.com/NTUST-ISC/ClubManageTool/assets/55608737/c6ce292a-8610-4833-b818-6a43ab4b9292)


### .xlsx
Fill in `member.xlsx`.  
(if a memeber is `社員`, its `職稱` column can be empty)  
(if member didn't provide phone number, `電話` column can be empty)
![image](https://github.com/NTUST-ISC/ClubManageTool/assets/55608737/f0e785d5-096e-4fc9-ab4c-5fc6cb209dfa)


### Run 
```shell
python main.py
```

