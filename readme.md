# 自動建立每月的時間表並生成EXCEL


##### 版本號: V1.2
 
## 功能特色:

1. 實時取得澳門假期

2. 自動把假期相對應的儲存格填上顏色

3. 自動生成每月的月份並放到對應的儲存格

4. 自動把日期對應星期幾

5. 自動生成 EXCEL並自動生成檔案名稱 

6. 自動把星期六和日相對應的儲存格填上顏色

7. 不頻繁向網站請求hoilday.isc 檔案 (備注1)

8. 用戶可從conf.ini 文件中自行更改想要的參數

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
1. 打開EXE version 資料夾裡面的exe檔
2. 輸入年份
3. 輸入月份
4. 看見根目錄下產生新的excel檔


## 根目錄文件
1. excel.py     (主程式)
2. origin.xlsx  (程式關連文件,更改文件會影響輸出格式)
3. conf.ini     (程式配置文件，用戶可根據需要更改設定)

#### 其他檔案(可刪除)
1. hoilday.ics  (從線上實時取得的澳門政府假期，運行程式一次會更新取代舊文件)
2. Daily Report of ...  .xlsx (最後結果輸出的檔案) 

## 備注
1. 為了不頻繁向網站請求hoilday.isc 檔案，而發生網站的防火牆禁止請求，程式會按照hoilday.ics 的修改日期來判斷檔案是否為當天的最新版, 若是當天的檔案,則不會重覆下載，反之hoilday.isc檔案不存在或非當天下載的檔案，則會從官網下載，
    1.1 若想每次執行程式都更新的話可在conf.ini 的文件中的day_diff 改成 0 . 
    1.2 若果想N天才向網站請求一次hoilday.isc，可自行更改conf.ini 的文件中的day_diff 的數字。
