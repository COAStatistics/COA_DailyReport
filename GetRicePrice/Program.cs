using System;
using CsQuery;
using System.Configuration;
using System.Net;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Globalization;

namespace GetRicePrice
{
    /// <summary>
    /// 取得米價
    /// 1.從目標網址取得當月份每天的米價並存入資料庫
    /// 2.如果資料庫中已存在該日的米價，就更新數值
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            //setup            
            DateTime startDate = new DateTime();
            DateTime endDate = new DateTime();
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
                startDate = new DateTime(DateTime.Now.AddMonths(-1).Year, DateTime.Now.AddMonths(-1).Month, 1);
                endDate = DateTime.Now;
            }
            Console.Write("糧價抓取日期：{0} -- {1}",startDate.ToString("yyyy/MM/01"), endDate.ToString("yyyy/MM/dd"));

            //資料庫連線
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    foreach (DateTime date in EachMonth(startDate, endDate))
                    {

                        //組合出目標URL以及查詢參數
                        string url =
                        ConfigurationManager.AppSettings["URL_RicePrice"] + "?" +
                        "city_name=" + "台北市" + "&" +
                        "city_id=" + "25" + "&" +
                        "year=" + date.ToString("yyyy", tc) + "&" +
                        "month=" + date.ToString("MM");

                        //對目標url發出request            
                        WebRequest myRequest = WebRequest.Create(url);
                        myRequest.Method = "GET";
                        //取得伺服器response
                        WebResponse myResponse = myRequest.GetResponse();
                        
                        //將response的內容讀出來
                        using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                        {                            
                            string result = sr.ReadToEnd();
                            //讀出來的內容是字串，使用CQ轉為HTML DOM物件
                            CQ dom = CQ.Create(result);
                            int trcount = dom["#result_table tr"].Length;

                            #region (1)日資料
                            comm.CommandText = @"
                                IF NOT EXISTS(
                                    SELECT * FROM rice_price
                                    WHERE year = @year AND month = @month AND days = @days
                                )
                                BEGIN
                                    INSERT INTO 
                                        rice_price(year,month,days,pr_1japt,pr_1tsait,pr_1sangt,pr_1glutrt,pr_1glutlt,pr_2japt,pr_2tsait,pr_2sangt,pr_2glutrt,pr_2glutlt,br_2japt,br_2tsait,br_2sangt,br_2glutrt,br_2glutlt,pa_japt,pa_tsait,pa_sangt,pa_glutrt,pa_glutlt,updatetime)
                                    VALUES(@year,@month,@days,@pr_1japt,@pr_1tsait,@pr_1sangt,@pr_1glutrt,@pr_1glutlt,@pr_2japt,@pr_2tsait,@pr_2sangt,@pr_2glutrt,@pr_2glutlt,@br_2japt,@br_2tsait,@br_2sangt,@br_2glutrt,@br_2glutlt,@pa_japt,@pa_tsait,@pa_sangt,@pa_glutrt,@pa_glutlt,GETDATE())
                                END
                                ELSE
                                UPDATE rice_price
                                SET 
                                    pr_1japt = @pr_1japt,
                                    pr_1tsait = @pr_1tsait,
                                    pr_1sangt = @pr_1sangt,
                                    pr_1glutrt = @pr_1glutrt,
                                    pr_1glutlt = @pr_1glutlt,

                                    pr_2japt = @pr_2japt,
                                    pr_2tsait = @pr_2tsait,
                                    pr_2sangt = @pr_2sangt,
                                    pr_2glutrt = @pr_2glutrt,
                                    pr_2glutlt = @pr_2glutlt,

                                    br_2japt = @br_2japt,
                                    br_2tsait = @br_2tsait,
                                    br_2sangt = @br_2sangt,
                                    br_2glutrt = @br_2glutrt,
                                    br_2glutlt = @br_2glutlt,

                                    pa_japt = @pa_japt,
                                    pa_tsait = @pa_tsait,
                                    pa_sangt = @pa_sangt,
                                    pa_glutrt = @pa_glutrt,
                                    pa_glutlt = @pa_glutlt,
                    
                                    updatetime = GETDATE()
                                WHERE year = @year AND month = @month AND days = @days
                            ";
                            //填入query parameter，並執行query
                            for (int i = 3; i < trcount - 1; i++)
                            {
                                comm.Parameters.Clear();
                                comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                                comm.Parameters.AddWithValue("@month", date.Month);
                                comm.Parameters.AddWithValue("@days", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(0)"].Html());

                                comm.Parameters.AddWithValue("@pr_1japt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(1)"].Html());
                                comm.Parameters.AddWithValue("@pr_1tsait", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(2)"].Html());
                                comm.Parameters.AddWithValue("@pr_1sangt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(3)"].Html());
                                comm.Parameters.AddWithValue("@pr_1glutrt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(4)"].Html());
                                comm.Parameters.AddWithValue("@pr_1glutlt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(5)"].Html());

                                comm.Parameters.AddWithValue("@pr_2japt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(6)"].Html());
                                comm.Parameters.AddWithValue("@pr_2tsait", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(7)"].Html());
                                comm.Parameters.AddWithValue("@pr_2sangt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(8)"].Html());
                                comm.Parameters.AddWithValue("@pr_2glutrt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(9)"].Html());
                                comm.Parameters.AddWithValue("@pr_2glutlt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(10)"].Html());

                                comm.Parameters.AddWithValue("@br_2japt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(11)"].Html());
                                comm.Parameters.AddWithValue("@br_2tsait", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(12)"].Html());
                                comm.Parameters.AddWithValue("@br_2sangt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(13)"].Html());
                                comm.Parameters.AddWithValue("@br_2glutrt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(14)"].Html());
                                comm.Parameters.AddWithValue("@br_2glutlt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(15)"].Html());

                                comm.Parameters.AddWithValue("@pa_japt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(16)"].Html());
                                comm.Parameters.AddWithValue("@pa_tsait", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(17)"].Html());
                                comm.Parameters.AddWithValue("@pa_sangt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(18)"].Html());
                                comm.Parameters.AddWithValue("@pa_glutrt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(19)"].Html());
                                comm.Parameters.AddWithValue("@pa_glutlt", dom["#result_table tr:eq(" + i.ToString() + ") td:eq(20)"].Html());
                                comm.Parameters.AddWithValue("@updatetime", DateTime.Now);
                                comm.ExecuteNonQuery();
                            }
                            #endregion

                            #region (2)月資料
                            comm.CommandText = @"
                                IF NOT EXISTS(
                                    SELECT * FROM monthly_rice
                                    WHERE year = @year AND month = @month
                                )
                                BEGIN
                                    INSERT INTO 
                                        monthly_rice(year,month,pr_1japt,pr_1tsait,pr_1sangt,pr_1glutrt,pr_1glutlt,pr_2japt,pr_2tsait,pr_2sangt,pr_2glutrt,pr_2glutlt,br_2japt,br_2tsait,br_2sangt,br_2glutrt,br_2glutlt,pa_japt,pa_tsait,pa_sangt,pa_glutrt,pa_glutlt,updatetime)
                                    VALUES(@year,@month,@pr_1japt,@pr_1tsait,@pr_1sangt,@pr_1glutrt,@pr_1glutlt,@pr_2japt,@pr_2tsait,@pr_2sangt,@pr_2glutrt,@pr_2glutlt,@br_2japt,@br_2tsait,@br_2sangt,@br_2glutrt,@br_2glutlt,@pa_japt,@pa_tsait,@pa_sangt,@pa_glutrt,@pa_glutlt,GETDATE())
                                END
                                ELSE
                                UPDATE monthly_rice
                                SET 
                                    pr_1japt = @pr_1japt,
                                    pr_1tsait = @pr_1tsait,
                                    pr_1sangt = @pr_1sangt,
                                    pr_1glutrt = @pr_1glutrt,
                                    pr_1glutlt = @pr_1glutlt,

                                    pr_2japt = @pr_2japt,
                                    pr_2tsait = @pr_2tsait,
                                    pr_2sangt = @pr_2sangt,
                                    pr_2glutrt = @pr_2glutrt,
                                    pr_2glutlt = @pr_2glutlt,

                                    br_2japt = @br_2japt,
                                    br_2tsait = @br_2tsait,
                                    br_2sangt = @br_2sangt,
                                    br_2glutrt = @br_2glutrt,
                                    br_2glutlt = @br_2glutlt,

                                    pa_japt = @pa_japt,
                                    pa_tsait = @pa_tsait,
                                    pa_sangt = @pa_sangt,
                                    pa_glutrt = @pa_glutrt,
                                    pa_glutlt = @pa_glutlt,
                    
                                    updatetime = GETDATE()
                                WHERE year = @year AND month = @month
                            ";
                            //填入query parameter，並執行query
                           
                            if (date.ToString("yyyyMM") != DateTime.Now.ToString("yyyyMM")) {

                                comm.Parameters.Clear();
                                comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                                comm.Parameters.AddWithValue("@month", date.Month);

                                comm.Parameters.AddWithValue("@pr_1japt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(1)"].Html());
                                comm.Parameters.AddWithValue("@pr_1tsait", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(2)"].Html());
                                comm.Parameters.AddWithValue("@pr_1sangt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(3)"].Html());
                                comm.Parameters.AddWithValue("@pr_1glutrt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(4)"].Html());
                                comm.Parameters.AddWithValue("@pr_1glutlt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(5)"].Html());

                                comm.Parameters.AddWithValue("@pr_2japt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(6)"].Html());
                                comm.Parameters.AddWithValue("@pr_2tsait", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(7)"].Html());
                                comm.Parameters.AddWithValue("@pr_2sangt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(8)"].Html());
                                comm.Parameters.AddWithValue("@pr_2glutrt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(9)"].Html());
                                comm.Parameters.AddWithValue("@pr_2glutlt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(10)"].Html());

                                comm.Parameters.AddWithValue("@br_2japt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(11)"].Html());
                                comm.Parameters.AddWithValue("@br_2tsait", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(12)"].Html());
                                comm.Parameters.AddWithValue("@br_2sangt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(13)"].Html());
                                comm.Parameters.AddWithValue("@br_2glutrt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(14)"].Html());
                                comm.Parameters.AddWithValue("@br_2glutlt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(15)"].Html());

                                comm.Parameters.AddWithValue("@pa_japt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(16)"].Html());
                                comm.Parameters.AddWithValue("@pa_tsait", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(17)"].Html());
                                comm.Parameters.AddWithValue("@pa_sangt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(18)"].Html());
                                comm.Parameters.AddWithValue("@pa_glutrt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(19)"].Html());
                                comm.Parameters.AddWithValue("@pa_glutlt", dom["#result_table tr:eq(" + (trcount - 1).ToString() + ") td:eq(20)"].Html());
                                comm.Parameters.AddWithValue("@updatetime", DateTime.Now);

                                try
                                {
                                    comm.ExecuteNonQuery();
                                }
                                catch {
                                    //
                                }                                
                            }

                            #endregion
                        }
                        myResponse.Close();

                    }
                }

            }            

        }
        //Generator
        public static IEnumerable<DateTime> EachMonth(DateTime from, DateTime thru)
        {
            for (var date = from.Date; date.Date <= thru.Date; date = date.AddMonths(1))
                yield return date;
        }

    }

}
