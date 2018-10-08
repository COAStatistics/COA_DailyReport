using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using CsQuery;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System;
using OpenQA.Selenium.Support.UI;
using System.Globalization;
using static OpenQA.Selenium.IJavaScriptExecutor;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace GetFlowerProduct
{
    class Program
    {
        static DateTime startDate { get; set; }
        static DateTime endDate { get; set; }

        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            //setup            
            if (args.Length != 0)
            {
                try
                {
                    var arg0 = Convert.ToDateTime(args[0]);
                    var arg1 = Convert.ToDateTime(args[1]);
                    if (arg1 <= DateTime.Now && arg0 <= arg1)
                    {
                        startDate = arg0;
                        endDate = arg1;
                    }
                }
                catch
                {
                    Console.WriteLine("FormatError.");
                }
            }
            else
            {
                startDate = DateTime.Now.AddDays(-14);
                endDate = DateTime.Now;
            }
            Console.WriteLine("花卉抓取日期：{0} -- {1}", startDate.ToString("yyyy/MM/dd"), endDate.ToString("yyyy/MM/dd"));
            Console.WriteLine("=============================== ===============================");

            //資料庫連線
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();
                    //取得設定值
                    comm.CommandText = @"
                        SELECT 
                            crops.id as cropid,
                            crops.code as cropcode,
                            config.name as configname
                        FROM
                            config,
                            crops
                        WHERE 
                            crops.configId = config.id
                        AND 
                            config.type = 'FlowerProduct'
                        AND
                            isTrack = 'Y'                   
                    ";
                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        sda.Fill(dt);
                    }

                    foreach (DataRow dr in dt.Rows)
                    {
                        int cropid = (int)dr["cropid"];
                        string configname = (string)dr["configname"];
                        string cropcode = dr.IsNull("cropcode") ? String.Empty : (string)dr["cropcode"];
                        comm.CommandText = @"
                            IF NOT EXISTS(
                                SELECT * FROM crops_price
                                WHERE cropid = @cropid
                                AND year = @year
                                AND month = @month
                                AND days = @days
                            )
                            BEGIN 
                                INSERT INTO 
                                    crops_price(cropid,year,month,days,avg,nt,updateTime)
                                VALUES(@cropid,@year,@month,@days,@avg,@nt,GETDATE())
                            END
                            ELSE
                            UPDATE crops_price
                            SET 
                                avg = @avg,nt = @nt,updateTime = GETDATE()
                            WHERE cropid = @cropid
                            AND year = @year
                            AND month = @month
                            AND days = @days      
                        ";

                        Stopwatch watch = new Stopwatch();
                        watch.Start();

                        List<MarketDailyTransObj> list = new List<MarketDailyTransObj>();
                        foreach (DateTime month in EachMonth(startDate, endDate))
                        {
                            var start = month.ToString("yyyy/MM/dd", tc);
                            var end = month.AddMonths(1) > endDate ? endDate.ToString("yyyy/MM/dd", tc) : month.AddMonths(1).ToString("yyyy/MM/dd", tc);

                            if (configname == "花卉")
                            {
                                Console.Write(String.Format("開啟模擬器抓取全部花卉資料...", configname));
                                list = GetAllFlowerData(start, end);
                                watch.Stop();
                                Console.WriteLine(String.Format("完成時間：{0}秒", watch.Elapsed.TotalSeconds));
                            }
                            else
                            {
                                Console.Write(String.Format("開啟模擬器抓取單一花卉資料...", configname));
                                list = GetSingleFlowerData(configname, cropcode, start, end);
                                watch.Stop();
                                Console.WriteLine(String.Format("完成時間：{0}秒", watch.Elapsed.TotalSeconds));
                            }

                            foreach (var obj in list)
                            {
                                comm.Parameters.Clear();
                                comm.Parameters.AddWithValue("@cropid", cropid);
                                comm.Parameters.AddWithValue("@year", obj.date.ToString("yyyy", tc));
                                comm.Parameters.AddWithValue("@month", obj.date.Month);
                                comm.Parameters.AddWithValue("@days", obj.date.Day);
                                comm.Parameters.AddWithValue("@avg", obj.tov_day / obj.tv_day);
                                comm.Parameters.AddWithValue("@nt", obj.tv_day);
                                comm.ExecuteNonQuery();
                            }
                        }

                    }
                }
            }
        }

        //Generator
        public static IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var date = from.Date; date.Date <= thru.Date; date = date.AddDays(1))
                yield return date;
        }

        public static IEnumerable<DateTime> EachMonth(DateTime from, DateTime thru) {
            for (var date = from.Date; date.Date <= thru.Date; date = date.AddMonths(1))
                yield return date;
        }



        public static List<MarketDailyTransObj> GetAllFlowerData(string startDate, string endDate) {

            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            string url = ConfigurationManager.AppSettings["URL_AllFlowerTrans"].ToString();
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
            var wait = new WebDriverWait(driver, TimeSpan.FromMinutes(2));            

            driver.Navigate().GoToUrl(url);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            //輸入全部市場    
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ctl00_contentPlaceHolder_txtMarket")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtMarket').innerHTML='全部市場'"));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_hfldMarketNo').setAttribute('value', '{0}')", "ALL"));

            //按下期間單選按鈕
            var radioBtn_term = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ctl00_contentPlaceHolder_ucDateScope_rblDateScope_1")));
            radioBtn_term.Click();

            //填入起始日期
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_txtSTransDate")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtSTransDate').setAttribute('value', '{0}')", startDate));
            //填入截止日期
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_txtETransDate")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtETransDate').setAttribute('value', '{0}')", endDate));

            //按下送出
            var submitBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ctl00_contentPlaceHolder_btnQuery")));
            submitBtn.Click();

            //取得結果表格
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_panel")));

            List<MarketDailyTransObj> list = new List<MarketDailyTransObj>();

            string table = (string)js.ExecuteScript("return document.getElementById('ctl00_contentPlaceHolder_panel').innerHTML");

            list = GetData(table, 2, 3, tc);

            driver.Close();
            driver.Quit();
            return list;

        }

        public static List<MarketDailyTransObj> GetSingleFlowerData(string configname, string cropcode,string startDate, string endDate)
        {
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            string url = ConfigurationManager.AppSettings["URL_SingleFlowerTrans"].ToString();
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
            var wait = new WebDriverWait(driver, TimeSpan.FromMinutes(2));

            driver.Navigate().GoToUrl(url);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            //輸入全部市場    
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ctl00_contentPlaceHolder_txtMarket")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtMarket').innerHTML='全部市場'"));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_hfldMarketNo').setAttribute('value', '{0}')", "ALL"));

            //輸入產品
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ctl00_contentPlaceHolder_txtProduct")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtProduct').innerHTML='{0} {1}'",cropcode, configname));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_hfldProductNo').setAttribute('value', '{0}')", cropcode));

            //按下期間單選按鈕
            var radioBtn_term = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ctl00_contentPlaceHolder_ucDateScope_rblDateScope_1")));
            radioBtn_term.Click();

            //填入起始日期
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_txtSTransDate")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtSTransDate').setAttribute('value', '{0}')", startDate));

            //填入截止日期
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_txtETransDate")));
            js.ExecuteScript(String.Format("document.getElementById('ctl00_contentPlaceHolder_txtETransDate').setAttribute('value', '{0}')", endDate));

            //按下送出
            var submitBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ctl00_contentPlaceHolder_btnQuery")));
            submitBtn.Click();

            //取得結果表格
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_contentPlaceHolder_panel")));

            List<MarketDailyTransObj> list = new List<MarketDailyTransObj>();

            string table = (string)js.ExecuteScript("return document.getElementById('ctl00_contentPlaceHolder_panel').innerHTML");

            list = GetData(table, 9, 7, tc);

            driver.Close();
            driver.Quit();
            return list;

        }

        public static List<MarketDailyTransObj> GetData(string table, int tv_day_col, int avg_day_col, CultureInfo tc) {

            CQ dom = CQ.Create(table);

            var trcount = dom["table:last-child tr"].Length;

            //用DataTable存取網頁上的表格資料，以反覆讀取
            DataTable dt = new DataTable();
            DataColumn column;
            DataRow row;

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "date";
            dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "tv_day";
            dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "tov_day";
            dt.Columns.Add(column);

            List<MarketDailyTransObj> list = new List<MarketDailyTransObj>();

            for (int i = 2; i < trcount; i++)
            {
                row = dt.NewRow();
                row["date"] = dom[String.Format("table:last-child tr:eq({0}) > td:eq(0)", i.ToString())].Html().Trim();
                var tv_day = Convert.ToDecimal(dom[String.Format("table:last-child tr:eq({0}) > td:eq({1})", i.ToString(), tv_day_col)].Text());
                var avg = Convert.ToDecimal(dom[String.Format("table:last-child tr:eq({0}) > td:eq({1})", i.ToString(), avg_day_col)].Text());
                row["tv_day"] = tv_day;
                row["tov_day"] = avg * tv_day;
                dt.Rows.Add(row);
            }

            foreach (DateTime date in EachDay(startDate, endDate))
            {
                MarketDailyTransObj obj = new MarketDailyTransObj();
                obj.date = date;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["date"].ToString() == date.ToString("yyyy/MM/dd", tc))
                    {

                        obj.tv_day += (decimal)dr["tv_day"];
                        obj.tov_day += (decimal)dr["tov_day"];

                    }
                }
                if (obj.tv_day != 0 && obj.tov_day != 0)
                {
                    list.Add(obj);
                }                
            }

            return list;

        }

    }

    public class MarketDailyTransObj {
        public DateTime date { get; set; }       
        public decimal tv_day { get; set; }//總交易量        
        public decimal tov_day { get; set; }//總交易金額
    }

}
