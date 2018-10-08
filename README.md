# 自動化產出日報表

	由數個不同網址抓取資料
	抓取回來的資料，會存入資料庫中，經過運算與排版後輸出為報表
	此方案包含數個抓取資料的程式，一個產出報表的程式及一個使用者介面
---
### 取得米價(GetRicePrice)
資料來源：
[行政院農業委員會農糧署 - 糧價查詢](http://210.69.25.143/report "http://210.69.25.143/report")

### 取得農作物交易價格(GetFarmProduct)
資料來源：
[行政院農業委員會農糧署 - 農產品批發市場交易行情站](http://amis.afa.gov.tw/main/Main.aspx "http://amis.afa.gov.tw/main/Main.aspx")

### 取得畜產交易價格(GetLiveStocksProduct)
資料來源：
[行政院農業委員會農糧署 - 畜產行情資訊網](http://ppg.naif.org.tw/naif/MarketInformation/Pig/TranStatistics.aspx")

### 取得花卉交易價格(GetFlowerProduct)
* 資料來源：
[行政院農業委員會農糧署 - 農產品批發市場交易行情站](http://amis.afa.gov.tw/main/Main.aspx "http://amis.afa.gov.tw/main/Main.aspx")
* 註解：
由於花卉細項眾多，為了更快速且準確取得全花卉、大項統計，採WebDriver方式抓取資料

~~### 取得產地月交易價格(GetLocalPriceMonthly)~~
資料來源：
[行政院農業委員會農糧署 - 農產品價格查報系統](http://apis.afa.gov.tw/pagepub/AppInquiryPage.aspx "http://apis.afa.gov.tw/pagepub/AppInquiryPage.aspx")

### 日報表(DailyReport)
複製範本檔，填入當日資訊，輸出為excel

### 日報表使用者介面(DailyReportGUI)
依序執行以上程式，產生日報表

### 資料庫(Database)
* 安裝SQL Server2008以上版本
* 還原dailyreport.bak