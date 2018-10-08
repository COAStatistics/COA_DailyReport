using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.Serialization;

namespace GetLocalPriceMonthly
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            //資料庫連線
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    //取得農作物設定值
                    comm.CommandText = @"
                        SELECT 
                            config.id as configid,
                            config.name as configname,
                            crops.name as cropname
                        FROM
                            config,
                            crops
                        WHERE 
							config.id = crops.configId
						AND
                            type = 'FarmProduct'
                        AND 
                            source = '產地價格'
                        AND
                            isTrack = 'Y'
                    ";
                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        sda.Fill(dt);
                    }

                    //對每個目標農作物，組出目標URL，取得資料
                    foreach (DataRow dr in dt.Rows)
                    {
                        int configid = (int)dr["configid"];

                        //加入計時器
                        Stopwatch watch = new Stopwatch();
                        watch.Start();
                        Console.Write(String.Format("發送{0}頁面請求...", dr["cropname"]));

                        //取得日期範圍:近兩個月
                        string url =
                            ConfigurationManager.AppSettings["URL_LocalPriceMonthly"] + "&$filter=作物+like+" + dr["cropname"].ToString();

                        WebRequest myRequest = WebRequest.Create(url);
                        myRequest.Method = "GET";
                        WebResponse myResponse = myRequest.GetResponse();

                        watch.Stop();
                        Console.WriteLine(String.Format("響應時間：{0}秒", watch.Elapsed.TotalSeconds));

                        using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                        {
                            string result = sr.ReadToEnd();
                            //將response的字串，反序列化為FarmTransObj物件
                            List<LocalPriceTransObj> json = JsonConvert.DeserializeObject<List<LocalPriceTransObj>>(result);

                            comm.CommandText = @"
                                IF NOT EXISTS(
                                    SELECT * FROM monthly_local
                                    WHERE configid = @configid
                                    AND year = @year
                                    AND month = @month
                                )
                                BEGIN 
                                    INSERT INTO 
                                        monthly_local(configid,year,month,avg,updateTime)
                                    VALUES(@configid,@year,@month,@avg,GETDATE())
                                END
                                ELSE
                                UPDATE monthly_local
                                SET 
                                    avg = @avg ,updateTime = GETDATE()
                                WHERE configid = @configid
                                AND year = @year
                                AND month = @month                   
                            ";
                            //填入query parameter，並執行query
                            foreach (LocalPriceTransObj lpto in json)
                            {
                                foreach (TransObject lp in lpto.monthlypricelist) {

                                    //當月無價格不進資料庫
                                    decimal j = 1;
                                    if (!decimal.TryParse(lp.avg, NumberStyles.Any, CultureInfo.InvariantCulture, out j))
                                    {
                                        continue;
                                    }
                                    comm.Parameters.Clear();
                                    comm.Parameters.AddWithValue("@configid", configid);
                                    comm.Parameters.AddWithValue("@year", lpto.年份);
                                    comm.Parameters.AddWithValue("@month", lp.month);
                                    comm.Parameters.AddWithValue("@avg", lp.avg);
                                    comm.ExecuteNonQuery();
                                }
                            }
                        }
                        myResponse.Close();
                    }
                }
            }
        }
    }


    class LocalPriceTransObj
    {
        [JsonProperty("年份")]
        public string 年份 { get; set; }
        [JsonProperty("1月價格")]
        public string 一月價格 { get; set; }
        [JsonProperty("2月價格")]
        public string 二月價格 { get; set; }
        [JsonProperty("3月價格")]
        public string 三月價格 { get; set; }
        [JsonProperty("4月價格")]
        public string 四月價格 { get; set; }
        [JsonProperty("5月價格")]
        public string 五月價格 { get; set; }
        [JsonProperty("6月價格")]
        public string 六月價格 { get; set; }
        [JsonProperty("7月價格")]
        public string 七月價格 { get; set; }
        [JsonProperty("8月價格")]
        public string 八月價格 { get; set; }
        [JsonProperty("9月價格")]
        public string 九月價格 { get; set; }
        [JsonProperty("10月價格")]
        public string 十月價格 { get; set; }
        [JsonProperty("11月價格")]
        public string 十一月價格 { get; set; }
        [JsonProperty("12月價格")]
        public string 十二月價格 { get; set; }
        public List<TransObject> monthlypricelist { get; set; }

        [OnDeserialized()]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            var year = Convert.ToInt32(年份.Replace("年", "").Trim()) - 1911;
            年份 = year.ToString();

            monthlypricelist = new List<TransObject>();
            monthlypricelist.Add(new TransObject() { month = 1, avg = 一月價格 });
            monthlypricelist.Add(new TransObject() { month = 2, avg = 二月價格 });
            monthlypricelist.Add(new TransObject() { month = 3, avg = 三月價格 });
            monthlypricelist.Add(new TransObject() { month = 4, avg = 四月價格 });
            monthlypricelist.Add(new TransObject() { month = 5, avg = 五月價格 });
            monthlypricelist.Add(new TransObject() { month = 6, avg = 六月價格 });
            monthlypricelist.Add(new TransObject() { month = 7, avg = 七月價格 });
            monthlypricelist.Add(new TransObject() { month = 8, avg = 八月價格 });
            monthlypricelist.Add(new TransObject() { month = 9, avg = 九月價格 });
            monthlypricelist.Add(new TransObject() { month = 10, avg = 十月價格 });
            monthlypricelist.Add(new TransObject() { month = 11, avg = 十一月價格 });
            monthlypricelist.Add(new TransObject() { month = 12, avg = 十二月價格 });
        }
    }

    class TransObject
    {     
        public int month { get; set; }
        public string avg { get; set; }
    }
}
