using System;
using System.IO;
using System.Net;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using CsQuery;
using System.Diagnostics;

namespace GetHogProducts
{
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
                startDate = DateTime.Now.AddDays(-14);
                endDate = DateTime.Now;
            }
            Console.WriteLine("畜產(肉牛除外)抓取日期：{0} -- {1}", startDate.ToString("yyyy/MM/dd"), endDate.ToString("yyyy/MM/dd"));
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
                            livestocks.id as livestockid,
                            config.name as livestockname,
                            market.name as marketname
                        FROM
                            config,
                            livestocks
                        LEFT JOIN market
                        ON market.id = livestocks.marketid
                        WHERE 
                            livestocks.configid = config.id
                        AND 
                            config.type = 'LivestockProduct'
                        AND
                            config.isTrack = 'Y'
                    ";
                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        sda.Fill(dt);
                    }

                    //組出目標URL，取得資料
                    foreach (DataRow dr in dt.Rows)
                    {
                        int livestockid = (int)dr["livestockid"];
                        string livestockname = dr["livestockname"].ToString();
                        string marketname = Convert.IsDBNull(dr["marketname"])? String.Empty : dr["marketname"].ToString();
                        string url = String.Empty;
                        string command = @"
                            IF NOT EXISTS(
                                SELECT * FROM livestocks_price
                                WHERE livestockid = @livestockid
                                AND year = @year
                                AND month = @month
                                AND days = @days
                            )
                            BEGIN 
                                INSERT INTO 
                                    livestocks_price(livestockid,year,month,days,avg,nt,kg,updateTime)
                                VALUES(@livestockid,@year,@month,@days,@avg,@nt,@kg,GETDATE())
                            END
                            ELSE
                            UPDATE livestocks_price
                            SET 
                                avg = @avg,
                                nt = @nt,
                                kg = @kg,
                                updateTime = GETDATE()
                            WHERE livestockid = @livestockid
                            AND year = @year
                            AND month = @month
                            AND days = @days
                        ";

                        Stopwatch watch = new Stopwatch();
                        watch.Start();
                        Console.Write(String.Format("發送{0}{1}頁面請求並計算...", marketname, livestockname));

                        if (livestockname == "肉牛")
                        {
                            //#region 7.肉牛
                            ////對目標網址發出請求
                            //url = ConfigurationManager.AppSettings["URL_BeefPrice"];
                            //WebRequest myRequest = WebRequest.Create(url);
                            //myRequest.Method = "GET";
                            //WebResponse myResponse = myRequest.GetResponse();
                            //using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                            //{
                            //    string result = sr.ReadToEnd();
                            //    CQ dom = CQ.Create(result);
                            //    string avg = dom[".Content .ScrollForm tr:eq(1)>td:eq(3)"].Text();
                            //    //設定日期範圍:這周日往前推兩個星期
                            //    int dateofweek = Convert.ToInt32(DateTime.Now.DayOfWeek.ToString("d"));
                            //    DateTime lastdayofweek = DateTime.Now.AddDays(6 - dateofweek);
                            //    foreach (DateTime date in EachDay(lastdayofweek.AddDays(-14), lastdayofweek))
                            //    {
                            //        comm.CommandText = command;
                            //        comm.Parameters.Clear();
                            //        comm.Parameters.AddWithValue("@livestockid", livestockid);
                            //        comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                            //        comm.Parameters.AddWithValue("@month", date.Month);
                            //        comm.Parameters.AddWithValue("@days", date.Day);
                            //        comm.Parameters.AddWithValue("@avg", avg);
                            //        comm.Parameters.AddWithValue("@kg", "0");
                            //        comm.Parameters.AddWithValue("@nt", "0");
                            //        comm.ExecuteNonQuery();
                            //    }
                            //}
                            //myResponse.Close();
                            //#endregion                            
                        }
                        else
                        {
                            #region 1.毛豬2.努比亞雜交閹公羊3.土番鴨4.紅羽土雞5.白肉雞6.雞蛋
                            //設定日期範圍:從今天往前推算兩個星期
                            foreach (DateTime date in EachDay(startDate, endDate))
                            {
                                switch (livestockname)
                                {
                                    case "毛豬":
                                        url = ConfigurationManager.AppSettings["URL_HogPrice"] + String.Format("?TransDate={0}&MarketName={1}", date.ToString("yyyyMMdd", tc),marketname);
                                        break;
                                    case "努比亞雜交閹公羊":
                                        url = ConfigurationManager.AppSettings["URL_MuttonPrice"] + String.Format("?$filter=productName+like+努比亞+and+transDate+like+{0}+and+shortName+like+{1}", date.ToString("yyyy/MM/dd"),marketname);
                                        break;
                                    case "土番鴨":
                                        url = ConfigurationManager.AppSettings["URL_DuckPrice"] + String.Format("?StartDate={0}&EndDate={0}", date.ToString("yyyy/MM/dd"));
                                        break;
                                    case "紅羽土雞":
                                        url = ConfigurationManager.AppSettings["URL_NativeChickenPrice"] + String.Format("?StartDate={0}&EndDate={0}", date.ToString("yyyy/MM/dd"));
                                        break;
                                    case "白肉雞":
                                    case "雞蛋":
                                        url = ConfigurationManager.AppSettings["URL_BoiledChickenPrice"] + String.Format("?StartDate={0}&EndDate={0}", date.ToString("yyyy/MM/dd"));
                                        break;
                                }
                                if (String.IsNullOrEmpty(url))
                                {
                                    break;
                                }
                                WebRequest myRequest = WebRequest.Create(url);
                                myRequest.Method = "GET";
                                WebResponse myResponse = myRequest.GetResponse();
                                using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                                {
                                    string result = sr.ReadToEnd();
                                    comm.CommandText = command;
                                    //將response的字串反序列化
                                    var json = new List<LiveStock>();
                                    try
                                    {
                                        switch (livestockname)
                                        {
                                            case "毛豬":
                                                var json1 = JsonConvert.DeserializeObject<List<HogTransObj>>(result);
                                                foreach (HogTransObj h in json1)
                                                {
                                                    json.Add(h);
                                                }
                                                break;
                                            case "努比亞雜交閹公羊":
                                                var json2 = JsonConvert.DeserializeObject<List<GoatTransObj>>(result);
                                                foreach (GoatTransObj g in json2)
                                                {
                                                    json.Add(g);
                                                }
                                                break;
                                            case "土番鴨":
                                                var json3 = JsonConvert.DeserializeObject<List<DuckTransObj>>(result);
                                                foreach (DuckTransObj g in json3)
                                                {
                                                    json.Add(g);
                                                }
                                                break;
                                            case "紅羽土雞":
                                                var json4 = JsonConvert.DeserializeObject<List<NativeChickenTransObj>>(result);
                                                foreach (NativeChickenTransObj g in json4)
                                                {
                                                    json.Add(g);
                                                }
                                                break;
                                            case "白肉雞":
                                                var json5 = JsonConvert.DeserializeObject<List<BoiledChickenTransObj>>(result);
                                                foreach (BoiledChickenTransObj g in json5)
                                                {
                                                    json.Add(g);
                                                }
                                                break;
                                            case "雞蛋":
                                                var json6 = JsonConvert.DeserializeObject<List<EggTransObj>>(result);
                                                foreach (EggTransObj g in json6)
                                                {
                                                    json.Add(g);
                                                }
                                                break;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }

                                    foreach (LiveStock l in json)
                                    {
                                        //檢查價格是否為0
                                        decimal j = 1;
                                        if (decimal.TryParse(l.平均價格, NumberStyles.Any, CultureInfo.InvariantCulture, out j))
                                        {
                                            if (j == 0)
                                            {
                                                continue;
                                            }
                                        }
                                        comm.Parameters.Clear();
                                        comm.Parameters.AddWithValue("@livestockid", livestockid);
                                        comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                                        comm.Parameters.AddWithValue("@month", date.Month);
                                        comm.Parameters.AddWithValue("@days", date.Day);
                                        comm.Parameters.AddWithValue("@avg", l.平均價格);
                                        comm.Parameters.AddWithValue("@nt", l.交易數量);
                                        comm.Parameters.AddWithValue("@kg", l.平均重量);
                                        comm.ExecuteNonQuery();
                                    }
                                }
                                myResponse.Close();
                            }
                            #endregion
                        }
                        watch.Stop();
                        Console.WriteLine(String.Format("完成時間：{0}", watch.Elapsed.TotalSeconds));
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
    }

    class LiveStock
    {
        public string 交易數量 { get; set; }
        public string 平均重量 { get; set; }
        public string 平均價格 { get; set; }
        public bool 是否有產量 { get; set; }
    }
    class HogTransObj : LiveStock
    {
        [JsonProperty("規格豬(75公斤以上)-頭數")]
        public string nt { get; set; }
        [JsonProperty("規格豬(75公斤以上)-平均價格")]
        public string avg { get; set; }
        [JsonProperty("市場名稱")]
        public string market { get; set; }
        [JsonProperty("規格豬(75公斤以上)-平均重量")]
        public string kg { get; set; }        

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = nt;
            平均重量 = kg;
            平均價格 = avg;
            是否有產量 = true;
        }
    }

    class GoatTransObj : LiveStock
    {
        [JsonProperty("quantity")]
        public string nt { get; set; }
        [JsonProperty("avgPrice")]
        public string avg { get; set; }
        [JsonProperty("shortName")]
        public string market { get; set; }
        [JsonProperty("avgWeight")]
        public string kg { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = nt;
            平均重量 = kg;
            平均價格 = avg;
            是否有產量 = true;
        }
    }

    class DuckTransObj : LiveStock
    {
        [JsonProperty("土番鴨(75天)")]
        public double avg { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = "0";
            平均重量 = "0";
            //每台斤價格轉每公斤價格
            平均價格 = Convert.ToString(avg / 0.6);
            是否有產量 = false;
        }
    }

    class NativeChickenTransObj : LiveStock
    {
        [JsonProperty("紅羽土雞北區")]
        public double avg_north { get; set; }
        [JsonProperty("紅羽土雞中區")]
        public double avg_midth { get; set; }
        [JsonProperty("紅羽土雞南區")]
        public double avg_south { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = "0";
            平均重量 = "0";
            //每台斤價格轉每公斤價格
            平均價格 = Convert.ToString((avg_north + avg_midth + avg_south) / 3 / 0.6);
            是否有產量 = false;
        }
    }

    class BoiledChickenTransObj : LiveStock
    {
        [JsonProperty("白肉雞(1.75-1.95Kg)")]
        public double avg { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = "0";
            平均重量 = "0";
            //每台斤價格轉每公斤價格
            平均價格 = Convert.ToString(avg / 0.6);
            是否有產量 = false;
        }

    }

    class EggTransObj : LiveStock
    {
        [JsonProperty("雞蛋(產地)")]
        public double avg { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            交易數量 = "0";
            平均重量 = "0";
            //每台斤價格轉每公斤價格
            平均價格 = Convert.ToString(avg / 0.6);
            是否有產量 = false;
        }
    }
}


