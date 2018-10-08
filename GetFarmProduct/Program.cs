using System;
using System.IO;
using System.Net;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Diagnostics;
namespace GetFarmProduct
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
            Console.WriteLine("農產品(不含花卉)抓取日期：{0} -- {1}", startDate.ToString("yyyy/MM/dd"), endDate.ToString("yyyy/MM/dd"));
            Console.WriteLine("=============================== ===============================");

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
                            crops.id as cropid,
                            crops.name as cropname,
                            crops.code as cropcode,
                            market.name as marketname
                        FROM
                            config,
                            crops,
                            market
                        WHERE 
                            crops.configId = config.id
                        AND 
                            crops.marketid = market.id
                        AND 
                            config.type = 'FarmProduct'
                        AND 
                            config.source = '批發價格'
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
                        int cropid = (int)dr["cropid"];
                        string cropcode = (string)dr["cropcode"];

                        //加入計時器
                        Stopwatch watch = new Stopwatch();
                        watch.Start();
                        Console.Write(String.Format("發送{0}頁面請求...", dr["cropname"]));

                        foreach (DateTime month in EachMonth(startDate, endDate)) {

                            var start = month.ToString("yyyy.MM.dd", tc);
                            var end = month.AddMonths(1) > endDate ? endDate.ToString("yyyy.MM.dd", tc) : month.AddMonths(1).ToString("yyyy.MM.dd", tc);

                            //取得日期範圍
                            string url =
                                ConfigurationManager.AppSettings["URL_FarmTrans"] + "?" +
                                "StartDate=" + start + "&" +
                                "EndDate=" + end + "&" +
                                "Market=" + dr["marketname"] + "&crop=" + dr["cropname"];

                            WebRequest myRequest = WebRequest.Create(url);
                            myRequest.Method = "GET";
                            WebResponse myResponse = myRequest.GetResponse();

                            watch.Stop();
                            Console.WriteLine(String.Format("響應時間：{0}秒", watch.Elapsed.TotalSeconds));

                            using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), true))
                            {
                                string result = sr.ReadToEnd();
                                //將response的字串，反序列化為FarmTransObj物件
                                List<FarmTransObj> json = JsonConvert.DeserializeObject<List<FarmTransObj>>(result);

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
                                        crops_price(cropid,year,month,days,hp,mp,lp,avg,nt,updateTime)
                                    VALUES(@cropid,@year,@month,@days,@hp,@mp,@lp,@avg,@nt,GETDATE())
                                END
                                ELSE
                                UPDATE crops_price
                                SET 
                                    hp = @hp,mp = @mp,lp = @lp,avg = @avg,nt = @nt,updateTime = GETDATE()
                                WHERE cropid = @cropid
                                AND year = @year
                                AND month = @month
                                AND days = @days                     
                            ";
                                //填入query parameter，並執行query
                                foreach (FarmTransObj f in json)
                                {
                                    //市場代號不符不進資料庫
                                    if (f.作物代號 != cropcode)
                                    {
                                        continue;
                                    }
                                    //當日休市資料不進資料庫
                                    decimal j = 1;
                                    if (decimal.TryParse(f.平均價, NumberStyles.Any, CultureInfo.InvariantCulture, out j))
                                    {
                                        if (j == 0)
                                        {
                                            continue;
                                        }
                                    }
                                    comm.Parameters.Clear();
                                    comm.Parameters.AddWithValue("@cropid", cropid);
                                    comm.Parameters.AddWithValue("@year", Convert.ToInt32(f.交易日期.Split('.')[0]));
                                    comm.Parameters.AddWithValue("@month", Convert.ToInt32(f.交易日期.Split('.')[1]));
                                    comm.Parameters.AddWithValue("@days", Convert.ToInt32(f.交易日期.Split('.')[2]));
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

        public static IEnumerable<DateTime> EachMonth(DateTime from, DateTime thru)
        {
            for (var date = from.Date; date.Date <= thru.Date; date = date.AddMonths(1))
                yield return date;
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
