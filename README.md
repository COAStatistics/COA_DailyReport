# 自動化產出日報表

	由數個不同網址抓取資料
	抓取回來的資料，會存入資料庫中，經過運算與排版後輸出為報表
	此方案包含數個抓取資料的程式，與一個產出報表的程式
---

## 目錄
* 撈資料程式
  * [取得米價](#取得米價getriceprice)
  * [取得農作物交易價格](#取得農作物交易價格getfarmproduct))
* 產出報表程式
  * [日報表](#日報表dailyreport)
* 資料庫
  * [PostgreSQL](#資料庫database)

---
### 取得米價(GetRicePrice)
資料來源：
[行政院農業委員會農糧署 - 糧價查詢](http://210.69.25.143/report "http://210.69.25.143/report")

### 取得農作物交易價格(GetFarmProduct)
資料來源：
[行政院農業委員會農糧署 - 農產品批發市場交易行情站](http://amis.afa.gov.tw/main/Main.aspx "http://amis.afa.gov.tw/main/Main.aspx")

### 日報表(DailyReport)
複製範本檔，填入當日資訊，輸出為excel。

### 資料庫(Database)
* 安裝PostgreSQL 9.5以上版本
* 還原coa_opendata.backup