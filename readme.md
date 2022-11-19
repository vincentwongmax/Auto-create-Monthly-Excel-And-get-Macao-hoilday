# 自動建立每月的時間表並生成EXCEL


##### 版本號: V1.0
 
## 功能特色:

1. 實時取得澳門假期

2. 自動把假期相對應的儲存格填上顏色

3. 自動生成每月的月份並放到對應的儲存格

4. 自動把日期對應星期幾

5. 自動生成 EXCEL並自動生成檔案名稱 

6. 自動把星期六和日相對應的儲存格填上顏色


## 使用方法

方法1 : 
1. 安裝python 
2. 使用pip安裝所有script 中的模組
3. 直接使用CMD 打開excel.py 
`或`使用以下`command` 編譯成`exe`檔
```
1. 打開`CMD` 輸入`command` 安裝 `pip install pyinstaller`
2. 打開`Powershell` 輸入 `command`  `pyinstaller --onefile -w 'excel.py'`

```
4. 輸入年份
5. 輸入月份
6. 看見根目錄下產生新的excel檔

方法2:
1. 打開exe檔
2. 輸入年份
3. 輸入月份
4. 看見根目錄下產生新的excel檔


## 根目錄文件
1. excel.py     (主程式)
2. origin.xlsx  (程式關連文件,更改文件會影響輸出格式)

#### 其他檔案(可刪除)
1. hoilday.ics  (從線上實時取得的澳門政府假期，運行程式一次會更新取代舊文件)
