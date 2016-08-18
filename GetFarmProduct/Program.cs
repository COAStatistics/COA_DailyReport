using System;
using System.IO;
using System.Net;
using Npgsql;
using System.Configuration;
using System.Data;
using System.Globalization;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace GetFarmProduct
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            //資料庫連線
            using (NpgsqlConnection conn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (NpgsqlCommand comm = new NpgsqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    //取得農作物設定值
                    comm.CommandText = @"
                        SELECT 
                            crops.id as cid,
                            crops.name as cropname,
                            market.id as mid,
                            market.name as marketname
                        FROM
                            config,
                            crops,
                            market
                        WHERE 
                            crops.cid = config.id
                            AND crops.mid = market.id
                    ";
                    using (NpgsqlDataAdapter sda = new NpgsqlDataAdapter(comm))
                    {
                        sda.Fill(dt);
                    }

                    //對每個目標農作物，組出目標URL，取得資料
                    foreach (DataRow dr in dt.Rows)
                    {
                        int cid = (int)dr["cid"];
                        int mid = (int)dr["mid"];
                        //取得日期範圍:前一個月的1號~今天
                        string url =
                            ConfigurationManager.AppSettings["URL_FarmTrans"] + "?" +
                            "StartDate=" + DateTime.Now.AddMonths(-1).ToString("yyyy.MM.01", tc) + "&" +
                            "EndDate=" + DateTime.Now.ToString("yyyy.MM.dd", tc) + "&" +
                            "Market=" + dr["marketname"] + "&crop=" + dr["cropname"];

                        WebRequest myRequest = WebRequest.Create(url);
                        myRequest.Method = "GET";
                        WebResponse myResponse = myRequest.GetResponse();
                        using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                        {
                            string result = sr.ReadToEnd();
                            //將response的字串，反序列化為FarmTransObj物件
                            List<FarmTransObj> json = JsonConvert.DeserializeObject<List<FarmTransObj>>(result);

                            comm.CommandText = @"
                                INSERT INTO 
                                    crops_price(cid,mid,year,month,days,hp,mp,lp,avg,nt)
                                VALUES(@cid,@mid,@year,@month,@days,@hp,@mp,@lp,@avg,@nt)
                                ON CONFLICT (cid, mid, year, month, days)
                                DO 
                                UPDATE SET
                                    hp = @hp,
                                    mp = @mp,
                                    lp = @lp,
                                    avg = @avg,
                                    nt = @nt,
                                    updatetime = now()
                            ";
                            //填入query parameter，並執行query
                            foreach (FarmTransObj f in json)
                            {
                                comm.Parameters.Clear();
                                comm.Parameters.AddWithValue("@cid", cid);
                                comm.Parameters.AddWithValue("@mid", mid);
                                comm.Parameters.AddWithValue("@year", f.交易日期.Split('.')[0]);
                                comm.Parameters.AddWithValue("@month", f.交易日期.Split('.')[1]);
                                comm.Parameters.AddWithValue("@days", f.交易日期.Split('.')[2]);
                                comm.Parameters.AddWithValue("@hp", f.上價);
                                comm.Parameters.AddWithValue("@mp", f.中價);
                                comm.Parameters.AddWithValue("@lp", f.下價);
                                comm.Parameters.AddWithValue("@avg", f.平均價);
                                comm.Parameters.AddWithValue("@nt", f.交易量);
                                comm.ExecuteNonQuery();
                            }
                        }
                        myResponse.Close();
                    }
                }
            }
        }
    }

    class FarmTransObj
    {
        public string 交易日期 { get; set; }
        public string 作物代號 { get; set; }
        public string 作物名稱 { get; set; }
        public string 市場代號 { get; set; }
        public string 市場名稱 { get; set; }
        public string 上價 { get; set; }
        public string 中價 { get; set; }
        public string 下價 { get; set; }
        public string 平均價 { get; set; }
        public string 交易量 { get; set; }
    }
}
