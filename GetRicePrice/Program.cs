using System;
using Npgsql;
using CsQuery;
using System.Configuration;
using System.Net;
using System.IO;


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
            //取得目前的民國年
            string NowYear = (DateTime.Now.Year - 1911).ToString();
            //string NowYear = "105";

            //取得目前月份
            string NowMonth = DateTime.Now.Month.ToString();
            //string NowMonth = "8";

            //組合出目標URL以及查詢參數
            string url =
                ConfigurationManager.AppSettings["URL_RicePrice"] + "?" +
                "city_name=" + "台北市" + "&" +
                "city_id=" + "25" + "&" +
                "year=" + NowYear + "&" +
                "month=" + NowMonth;

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

                //資料庫連線
                using (NpgsqlConnection conn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
                {
                    using (NpgsqlCommand comm = new NpgsqlCommand())
                    {
                        comm.Connection = conn;
                        conn.Open();
                        comm.CommandText = @"
                            INSERT INTO 
                                rice_price_avg(year,month,days,pr_1japt,pr_1tsait,pr_1sangt,pr_1glutrt,pr_1glutlt,pr_2japt,pr_2tsait,pr_2sangt,pr_2glutrt,pr_2glutlt,br_2japt,br_2tsait,br_2sangt,br_2glutrt,br_2glutlt,pa_japt,pa_tsait,pa_sangt,pa_glutrt,pa_glutlt)
                            VALUES(@year,@month,@days,@pr_1japt,@pr_1tsait,@pr_1sangt,@pr_1glutrt,@pr_1glutlt,@pr_2japt,@pr_2tsait,@pr_2sangt,@pr_2glutrt,@pr_2glutlt,@br_2japt,@br_2tsait,@br_2sangt,@br_2glutrt,@br_2glutlt,@pa_japt,@pa_tsait,@pa_sangt,@pa_glutrt,@pa_glutlt)
                            ON CONFLICT (year, month, days)
                            DO 
                            UPDATE SET 
                                pr_1japt = EXCLUDED.pr_1japt,
                                pr_1tsait = EXCLUDED.pr_1tsait,
                                pr_1sangt = EXCLUDED.pr_1sangt,
                                pr_1glutrt = EXCLUDED.pr_1glutrt,
                                pr_1glutlt = EXCLUDED.pr_1glutlt,

                                pr_2japt = EXCLUDED.pr_2japt,
                                pr_2tsait = EXCLUDED.pr_2tsait,
                                pr_2sangt = EXCLUDED.pr_2sangt,
                                pr_2glutrt = EXCLUDED.pr_2glutrt,
                                pr_2glutlt = EXCLUDED.pr_2glutlt,

                                br_2japt = EXCLUDED.br_2japt,
                                br_2tsait = EXCLUDED.br_2tsait,
                                br_2sangt = EXCLUDED.br_2sangt,
                                br_2glutrt = EXCLUDED.br_2glutrt,
                                br_2glutlt = EXCLUDED.br_2glutlt,

                                pa_japt = EXCLUDED.pa_japt,
                                pa_tsait = EXCLUDED.pa_tsait,
                                pa_sangt = EXCLUDED.pa_sangt,
                                pa_glutrt = EXCLUDED.pa_glutrt,
                                pa_glutlt = EXCLUDED.pa_glutlt,
                    
                                updatetime = now()
                        ";

                        //填入query parameter，並執行query
                        for (int i = 3; i < trcount - 1; i++)
                        {
                            comm.Parameters.Clear();
                            comm.Parameters.AddWithValue("@year", NowYear);
                            comm.Parameters.AddWithValue("@month", NowMonth);
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
                    }
                }
            }

            //關閉
            myResponse.Close();
        }
    }
}
