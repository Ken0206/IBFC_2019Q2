### Step10.Copy Font.bat
```
刪除 D:\造字
複製新造字於 D:\造字
```
---
### Step13.OutlookConfig.vbs
```
outlook 2007, 2010, 2016
關閉自動下載 HTML 郵件中的圖片︰registry 可以設定
純文字模式讀取郵件︰registry 可以設定
關閉自動預覽郵件功能︰設定值存在 .pst 檔，目前只能執行 outlook 來設定

```
---
### Step01.SetupWorkPath_New.vbs
```
原 Step01.SetupWorkPath.vbs 使用方式︰
1. workpath.txt 必須存在，並清空內容
2. logpath.txt 必須存在，並手動設定路徑於此

新 Step01.SetupWorkPath_New.vbs 使用方式︰
1. 有確認是否執行
2. workpath.txt, logpath.txt 不須存在
3. 自動套用目前工作目錄，產生 logpath.txt，
4. logpath.txt 如果存在，則複寫

看新的是否合用
```   
