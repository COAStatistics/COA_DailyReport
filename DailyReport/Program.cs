using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using System.Configuration;
using NPOI.XSSF.UserModel;
using Npgsql;
using System.Globalization;
using System.Data;

namespace DailyReport
{
    class Program
    {
        static void Main(string[] args)
        {
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            #region 讀取範本excel檔中的資料
            XSSFWorkbook workbook;
            using (FileStream fs = new FileStream("日報表範本.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                workbook = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
            #endregion

            DataTable dt = new DataTable();

            using (NpgsqlConnection conn = new NpgsqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (NpgsqlCommand comm = new NpgsqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    #region 1.米價
                    int RowNum_稉種稻穀 = 8;
                    int RowNum_長糯白米 = 9;
                    double 稉種稻穀本週加總價格 = 0;
                    double 長糯白米躉售本週加總價格 = 0;
                    double 稉種稻穀上週加總價格 = 0;
                    double 長糯白米躉售上週加總價格 = 0;

                    comm.CommandText = @"
                        SELECT 
                            (rice_price_avg.pa_japt::numeric / 100::numeric)::numeric(8,4) AS field01,
                            (pr_2glutlt::numeric / 100::numeric)::numeric(8,4) AS field02                            
                        FROM 
                            rice_price_avg
                        WHERE 
                            year = @year AND month = @month AND days = @days
                    ";

                    for (int i = 0; i < 14; i++)
                    {
                        comm.Parameters.Clear();
                        DateTime targetDate = DateTime.Now.AddDays(i * -1);
                        comm.Parameters.AddWithValue("@year", targetDate.ToString("yyyy", tc));
                        comm.Parameters.AddWithValue("@month", targetDate.Month.ToString());
                        comm.Parameters.AddWithValue("@days", targetDate.Day.ToString());
                        using (NpgsqlDataAdapter sda = new NpgsqlDataAdapter(comm))
                        {
                            dt.Clear();
                            sda.Fill(dt);
                            sheet.GetRow(7).GetCell(18 - i).SetCellValue(targetDate.Month.ToString() + "月" + Environment.NewLine + targetDate.Day.ToString() + "日");

                            if (dt.Rows.Count > 0)
                            {
                                if (i < 7)
                                {
                                    double pa_japt = double.Parse(dt.Rows[0]["field01"].ToString());
                                    sheet.GetRow(RowNum_稉種稻穀).GetCell(18 - i).SetCellValue(pa_japt);
                                    稉種稻穀本週加總價格 += pa_japt;

                                    double pr_2glutlt = double.Parse(dt.Rows[0]["field02"].ToString());
                                    sheet.GetRow(RowNum_長糯白米).GetCell(18 - i).SetCellValue(pr_2glutlt);
                                    長糯白米躉售本週加總價格 += pr_2glutlt;
                                }
                                else
                                {
                                    double pa_japt = double.Parse(dt.Rows[0]["field01"].ToString());
                                    稉種稻穀上週加總價格 += pa_japt;
                                    double pr_2glutlt = double.Parse(dt.Rows[0]["field02"].ToString());
                                    長糯白米躉售上週加總價格 += pr_2glutlt;
                                }
                            }
                            else
                            {
                                //米價應該沒有休市日吧
                            }
                        }
                    }
                    double 稉種稻穀本週平均價格 = 稉種稻穀本週加總價格 / 7;
                    double 長糯白米躉售本週平均價格 = 長糯白米躉售本週加總價格 / 7;
                    double 稉種稻穀上週平均價格 = 稉種稻穀上週加總價格 / 7;
                    double 長糯白米躉售上週平均價格 = 長糯白米躉售上週加總價格 / 7;
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(7).SetCellValue(稉種稻穀本週平均價格);
                    sheet.GetRow(RowNum_長糯白米).GetCell(7).SetCellValue(長糯白米躉售本週平均價格);
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(22).SetCellValue(稉種稻穀上週平均價格);
                    sheet.GetRow(RowNum_長糯白米).GetCell(22).SetCellValue(長糯白米躉售上週平均價格);
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(11).SetCellValue(((稉種稻穀本週平均價格 - 稉種稻穀上週平均價格) / 稉種稻穀本週平均價格) * 100);
                    sheet.GetRow(RowNum_長糯白米).GetCell(11).SetCellValue(((長糯白米躉售本週平均價格 - 長糯白米躉售上週平均價格) / 長糯白米躉售本週平均價格) * 100);
                    #endregion

                    #region 2.農產品
                    comm.CommandText = @"
                        SELECT
                            id,
                            name, 
                            rownum
                        FROM
                            config
                        WHERE
                            track = true
                    ";
                    using (NpgsqlDataAdapter sda = new NpgsqlDataAdapter(comm))
                    {
                        dt.Clear();
                        sda.Fill(dt);
                    }
                    List<FarmProduct> farmProductList = new List<FarmProduct>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        FarmProduct farmProduct = new FarmProduct();
                        farmProduct.id = (int)dr["id"];
                        farmProduct.name = dr["name"].ToString();
                        farmProduct.rownum = (int)dr["rownum"] - 1;
                        farmProductList.Add(farmProduct);
                    }

                    foreach (FarmProduct farmProduct in farmProductList)
                    {
                        comm.Parameters.Clear();
                        StringBuilder sb = new StringBuilder();
                        sb.Append(@"
                            SELECT
                                config.name as cropname,
                                market.name as marketname,
                                crops.name as cropname_detail,
                                crops.code as cropcode,
                                crops_price.year || '.' || crops_price.month || '.' || crops_price.days as times,
                                crops_price.avg,
                                crops_price.nt
                            FROM
                                config,
                                market,
                                crops,
                                crops_price
                            WHERE
                                config.id = crops.cid
                                AND market.id = crops.mid
                                AND crops.id = crops_price.cid                                
                                AND config.id = @cropid
                                AND crops_price.year || '.' || crops_price.month || '.' || crops_price.days in (
                        ");
                        for (var i = 0; i < 14; i++)
                        {
                            sb.Append("@date" + i.ToString() + ",");
                            comm.Parameters.AddWithValue("@date" + i.ToString(), DateTime.Now.AddDays(i * -1).ToString("yyyy.MM.dd", tc));
                        }
                        sb.Length--;
                        sb.Append(")");
                        comm.Parameters.AddWithValue("@cropid", farmProduct.id);
                        comm.CommandText = sb.ToString();
                        using (NpgsqlDataAdapter sda = new NpgsqlDataAdapter(comm))
                        {
                            dt.Clear();
                            sda.Fill(dt);
                        }

                        foreach (DataRow dr in dt.Rows)
                        {
                            int daydiff = (int)Math.Floor((DateTime.Now - DateTime.ParseExact(dr["times"].ToString(), "yyyy.MM.dd", tc)).TotalDays);
                            farmProduct.tv_day[daydiff] += Double.Parse(dr["nt"].ToString());
                            farmProduct.tov_day[daydiff] += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["avg"].ToString());
                        }

                        for (int i = 0; i < 14; i++)
                        {
                            if (i < 7)
                            {
                                if (HasValue(farmProduct.tov_day[i] / farmProduct.tv_day[i]))
                                {
                                    farmProduct.tov_this += farmProduct.tov_day[i];
                                    farmProduct.tv_this += farmProduct.tv_day[i];
                                    sheet.GetRow(farmProduct.rownum).GetCell(18 - i).SetCellValue(farmProduct.tov_day[i] / farmProduct.tv_day[i]);
                                }
                                else
                                {
                                    sheet.GetRow(farmProduct.rownum).GetCell(18 - i).SetCellValue("-");
                                }
                            }
                            else
                            {
                                if (HasValue(farmProduct.tov_day[i] / farmProduct.tv_day[i]))
                                {
                                    farmProduct.tov_last += farmProduct.tov_day[i];
                                    farmProduct.tv_last += farmProduct.tv_day[i];
                                }
                                else
                                {

                                }
                            }
                        }
                        sheet.GetRow(farmProduct.rownum).GetCell(19).SetCellValue(farmProduct.tv_this);
                        sheet.GetRow(farmProduct.rownum).GetCell(20).SetCellValue(100 * (farmProduct.tv_this - farmProduct.tv_last) / farmProduct.tv_this);
                        sheet.GetRow(farmProduct.rownum).GetCell(7).SetCellValue(farmProduct.tov_this / farmProduct.tv_this);
                        sheet.GetRow(farmProduct.rownum).GetCell(22).SetCellValue(farmProduct.tov_last / farmProduct.tv_last);
                        sheet.GetRow(farmProduct.rownum).GetCell(11).SetCellValue(100 * ((farmProduct.tov_this / farmProduct.tv_this) - (farmProduct.tov_last / farmProduct.tv_last)) / (farmProduct.tov_this / farmProduct.tv_this));
                    }
                    
                    #endregion
                }
            }


            #region 設定excel檔名並寫入
            DateTime startDate = DateTime.Now.AddDays(-6);
            DateTime endDate = DateTime.Now;
            string dayofweek = "";
            switch (DateTime.Now.DayOfWeek.ToString("d"))
            {
                case "0":
                    dayofweek = "週日";
                    break;
                case "1":
                    dayofweek = "週一";
                    break;
                case "2":
                    dayofweek = "週二";
                    break;
                case "3":
                    dayofweek = "週三";
                    break;
                case "4":
                    dayofweek = "週四";
                    break;
                case "5":
                    dayofweek = "週五";
                    break;
                case "6":
                    dayofweek = "週六";
                    break;
            }
            using (FileStream fs = File.Open((startDate.Year - 1911).ToString() + "." + startDate.ToString("MM.dd") + "-" + (endDate.Year - 1911).ToString() + "." + endDate.ToString("MM.dd") + "_" + (endDate.Year - 1911).ToString() + "價格" + dayofweek + ".xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(fs);
                fs.Close();
            }
            #endregion
        }

        private static bool HasValue(double value)
        {
            return !Double.IsNaN(value) && !Double.IsInfinity(value);
        }
    }

    class FarmProduct
    {
        public int id { get; set; }
        public string name { get; set; }
        public int rownum { get; set; }//在excel中的row編號
        public double tv_this { get; set; }//本週總交易量
        public double tov_this { get; set; }//本週總成交金額
        public double tv_last { get; set; }//上週總交易量
        public double tov_last { get; set; }//上週總成交金額
        public double[] tv_day = new double[14];//當日交易量
        public double[] tov_day = new double[14];//當日交易金額
    }
}
