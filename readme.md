# 自動建立每月的時間表並生成EXCEL


##### 版本號: V1.3
 
## 功能特色:

1. 實時取得澳門假期

2. 自動把假期相對應的儲存格填上顏色

3. 自動生成每月的月份並放到對應的儲存格

4. 自動把日期對應星期幾

5. 自動生成 EXCEL並自動生成檔案名稱 

6. 自動把星期六和日相對應的儲存格填上顏色

7. 不頻繁向網站請求hoilday.isc 檔案 (備注1)

8. 用戶可從conf.ini 文件中自行更改想要的參數

---
## 使用方法

`方法1 `: 
1. 安裝python 
2. 使用pip安裝所有script 中的模組
3. 直接使用CMD 打開excel.py 
`或`使用以下`command` 編譯成`exe`檔
```
1. 打開`CMD` 輸入`command` 安裝 `pip install pyinstaller`
2. 打開`Powershell` 輸入 `command`  `pyinstaller --onefile -w 'excel.py'`

```
4. 大量輸出模式 
>已打開大量輸出模式
>    設定參數正確，看見根目錄下產生多個新的excel檔

>已關閉大量輸出模式
>   輸入年份
>   輸入月份
>   看見根目錄下產生新的excel檔

---
`方法2`:
1. 打開EXE version 資料夾裡面的exe檔
2. 看需要修改conf.ini 的文件
3. 大量輸出模式 
>已打開大量輸出模式
>    3.1 設定參數正確，看見根目錄下產生多個新的excel檔

>已關閉大量輸出模式
>   3.1 輸入年份
>   3.2 輸入月份
>   3.3 看見根目錄下產生新的excel檔

---
## 根目錄文件
1. excel.py     (主程式)
2. origin.xlsx  (程式關連文件,更改文件會影響輸出格式)
3. conf.ini     (程式配置文件，用戶可根據需要更改設定)

#### 其他檔案(可刪除)
1. hoilday.ics  (從線上實時取得的澳門政府假期，運行程式一次會更新取代舊文件)
2. Daily Report of ...  .xlsx (最後結果輸出的檔案) 

#### config文件說明

|  變數   | 說明  |
|  ----  | ----  |
| day_diff=1 | hoilday.isc 過了多少天才會向網站下載新的版本 注1|
| weekend_color=FCE4D6  | 設定週六和週日儲存格的顏色，顏色參數請參考HTML |
| hoilday_color=FFFF99  | 設定澳門假期儲存格的顏色，顏色參數請參考HTML |
| save_excel_name  | 輸出EXCEL 的名稱 |
| year_month_day_who_first=day  | 想要年月日,哪一個排隊,由左到由 day/month/year|
| year_month_day_who_second=month  | 想要年月日,哪一個排隊,由左到由 day/month/year|
| year_month_day_who_third=year  | 想要年月日,哪一個排隊,由左到由 day/month/year|
| excel_Formula_mode=false  | 對應的星期六日是否用excel 公式|

|  大批輸出模式   | 說明  |
|  ----  | ----  |
| muti_mode_on=true | 是否打開大批輸出模式,否果沒有打開(false)，則下面的參數不會生效|
| export_year | 輸入大批輸出的年份|
| export_start_month | 從哪個月份開始|
| export_end_month | 從哪個月份結束|


## 備注
1. 為了不頻繁向網站請求hoilday.isc 檔案，而發生網站的防火牆禁止請求，程式會按照hoilday.ics 的修改日期來判斷檔案是否為當天的最新版, 若是當天的檔案,則不會重覆下載，反之hoilday.isc檔案不存在或非當天下載的檔案，則會從官網下載，
    1.1 若想每次執行程式都更新的話可在conf.ini 的文件中的day_diff 改成 0 . 
    1.2 若果想N天才向網站請求一次hoilday.isc，可自行更改conf.ini 的文件中的day_diff 的數字。
