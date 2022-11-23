import os,time
from datetime import datetime

# 要檢查的檔案路徑
filepath = "hoilday.ics"

# 檢查檔案是否存在
if os.path.isfile(filepath):
  print("檔案存在。")
else:
  print("檔案不存在。")


modifiedTime = time.localtime(os.stat(filepath).st_mtime)  ##出來的是一串奇怪的時間格式
createdTime = time.localtime(os.stat(filepath).st_ctime)    ##出來的是一串奇怪的時間格式

# mTime = time.strftime('%Y-%m-%d %H:%M:%S', modifiedTime)
# cTime = time.strftime('%Y-%m-%d %H:%M:%S', createdTime)

mTimes = time.strftime('%Y-%m-%d', modifiedTime)   # 用time.strftime 來定義格式
cTimes = time.strftime('%Y-%m-%d', createdTime)     # 用time.strftime 來定義格式

# datetime.today()  ## 2022-11-23 13:15:20.309244
today = datetime.today().strftime("%Y-%m-%d")      # 用time.strftime 來定義格式,'%Y-%m-%d %H:%M:%S'
# print("modifiedTime " + mTime)
# print("createdTime " + cTime)

def demo(day1, day2):  ## 兩個時間相減等出日期
    time_array1 = time.strptime(day1, "%Y-%m-%d")     ## 先統一時間格式
    timestamp_day1 = int(time.mktime(time_array1))    ## 以秒數形式回傳，再轉做int 
    time_array2 = time.strptime(day2, "%Y-%m-%d")     ## 先統一時間格式
    timestamp_day2 = int(time.mktime(time_array2))    ## 以秒數形式回傳，再轉做int 
    result = (timestamp_day2 - timestamp_day1) // 60 // 60 // 24   ## // = 取整數
    return result

# day1 = "2018-07-09"
# day2 = "2020-09-26"

day_diff = demo(mTimes, today)
print("两个日期的间隔天数：{} ".format(day_diff))










