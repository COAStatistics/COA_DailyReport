using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Configuration;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace DailyReport
{
    class Program
    {
        static void Main(string[] args)
        {
            //產出日報表日期
            DateTime basedate = DateTime.Now.AddDays(-1);
            //產出日報表位置
            string path = String.Empty;
            //顯示清單
            string displayname = String.Empty;

            if (args.Length != 0)
            {
                try
                {
                    if (Convert.ToDateTime(args[0]) <= DateTime.Now) {
                        basedate = Convert.ToDateTime(args[0]);
                    }
                }
                catch {
                    Console.WriteLine("FormatError.");
                }
            }

            Console.WriteLine("產生日報表日期：" + basedate.ToString("yyyy年MM月dd日"));

            if (args.Length > 1) {
                path = Path.GetFullPath(args[1]);
                Console.WriteLine("產生日報表位置：" + args[1]);
            }

            if (args.Length > 2)
            {
                displayname = args[2];
                Console.WriteLine("顯示清單名稱：" + args[2]);
            }
            
            CultureInfo tc = new CultureInfo("zh-TW");
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();

            #region 讀取範本excel檔中的資料
            XSSFWorkbook workbook;
            using (FileStream fs = new FileStream("template.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                workbook = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
            #endregion

            DataTable dt = new DataTable();

            using (SqlConnection conn= new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    sheet.GetRow(5).GetCell(6).SetCellValue(String.Format("{0}\n年{1}月\n平均價格",basedate.AddYears(-1).ToString("yyyy",tc),basedate.Month.ToString()));

                    #region 1.米價
                    int RowNum_稉種稻穀 = 8;
                    int RowNum_長糯白米 = 9;
                    double 稉種稻穀本週加總價格 = 0;
                    double 長糯白米躉售本週加總價格 = 0;
                    double 稉種稻穀上週加總價格 = 0;
                    double 長糯白米躉售上週加總價格 = 0;
                    double 稉種稻穀去年同月平均價格 = 0;
                    double 長糯白米去年同月平均價格 = 0;

                    #region (1)取得近兩周金額及交易量
                    comm.CommandText = @"
                        SELECT 
                            (pa_japt / CAST(100 as float)) AS field01,
                            (pr_2glutlt / CAST(100 as float)) AS field02                            
                        FROM 
                            rice_price
                        WHERE 
                            year = @year AND month = @month AND days = @days
                    ";



                    for (int i = 0; i < 14; i++)
                    {
                        comm.Parameters.Clear();
                        DateTime targetDate = basedate.AddDays(i * -1);
                        comm.Parameters.AddWithValue("@year", targetDate.ToString("yyyy", tc));
                        comm.Parameters.AddWithValue("@month", targetDate.Month);
                        comm.Parameters.AddWithValue("@days", targetDate.Day);

                        if (i < 7) {
                            sheet.GetRow(7).GetCell(18 - i).SetCellValue(targetDate.Month.ToString() + "月" + Environment.NewLine + targetDate.Day.ToString() + "日");
                            sheet.GetRow(7).GetCell(29 - i).SetCellValue(targetDate.Month.ToString() + "月" + Environment.NewLine + targetDate.Day.ToString() + "日");
                        }                      

                        using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                        {
                            dt.Clear();
                            sda.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                if (i < 7)
                                {
                                    double pa_japt = double.Parse(dt.Rows[0]["field01"].ToString());
                                    sheet.GetRow(RowNum_稉種稻穀).GetCell(18 - i).SetCellValue(pa_japt);
                                    sheet.GetRow(RowNum_稉種稻穀).GetCell(29 - i).SetCellValue("-");
                                    稉種稻穀本週加總價格 += pa_japt;

                                    double pr_2glutlt = double.Parse(dt.Rows[0]["field02"].ToString());
                                    sheet.GetRow(RowNum_長糯白米).GetCell(18 - i).SetCellValue(pr_2glutlt);
                                    sheet.GetRow(RowNum_長糯白米).GetCell(29 - i).SetCellValue("-");
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
                                sheet.GetRow(RowNum_稉種稻穀).GetCell(18 - i).SetCellValue("-");
                                sheet.GetRow(RowNum_長糯白米).GetCell(18 - i).SetCellValue("-");
                                sheet.GetRow(RowNum_稉種稻穀).GetCell(29 - i).SetCellValue("-");
                                sheet.GetRow(RowNum_長糯白米).GetCell(29 - i).SetCellValue("-");
                            }
                        }
                    }
                    #endregion

                    #region (2)取得去年同月平均價格
                    comm.CommandText = @"
                        SELECT 
                            pa_japt / CAST(100 as float) AS field01,
                            pr_2glutlt / CAST(100 as float) AS field02                      
                        FROM 
                            monthly_rice
                        WHERE 
                            year = @year AND month = @month 
                    ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@year", basedate.AddYears(-1).ToString("yyyy", tc));
                    comm.Parameters.AddWithValue("@month", basedate.AddYears(-1).Month);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Clear();
                        sda.Fill(dt);

                        if (dt.Rows.Count > 0) {
                            稉種稻穀去年同月平均價格 = double.Parse(dt.Rows[0]["field01"].ToString());
                            長糯白米去年同月平均價格 = double.Parse(dt.Rows[0]["field02"].ToString());
                        }                        
                    }
                    #endregion

                    double 稉種稻穀本週平均價格 = 稉種稻穀本週加總價格 / 7;
                    double 長糯白米躉售本週平均價格 = 長糯白米躉售本週加總價格 / 7;
                    double 稉種稻穀上週平均價格 = 稉種稻穀上週加總價格 / 7;
                    double 長糯白米躉售上週平均價格 = 長糯白米躉售上週加總價格 / 7;
                    sheet.GetRow(7).GetCell(22).SetCellValue(String.Format("{0}~{1}",basedate.AddDays(-13).ToString("M/d"),basedate.AddDays(-7).ToString("M/d")));
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(6).SetCellValue(稉種稻穀去年同月平均價格);
                    sheet.GetRow(RowNum_長糯白米).GetCell(6).SetCellValue(長糯白米去年同月平均價格);
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(7).SetCellValue(稉種稻穀本週平均價格);
                    sheet.GetRow(RowNum_長糯白米).GetCell(7).SetCellValue(長糯白米躉售本週平均價格);
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(22).SetCellValue(稉種稻穀上週平均價格);
                    sheet.GetRow(RowNum_長糯白米).GetCell(22).SetCellValue(長糯白米躉售上週平均價格);
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(11).SetCellValue(((稉種稻穀本週平均價格 - 稉種稻穀上週平均價格) / 稉種稻穀本週平均價格) * 100);
                    sheet.GetRow(RowNum_長糯白米).GetCell(11).SetCellValue(((長糯白米躉售本週平均價格 - 長糯白米躉售上週平均價格) / 長糯白米躉售本週平均價格) * 100);
                    sheet.GetRow(RowNum_長糯白米).GetCell(19).SetCellValue("…");
                    sheet.GetRow(RowNum_長糯白米).GetCell(20).SetCellValue("-");
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(19).SetCellValue("…");
                    sheet.GetRow(RowNum_稉種稻穀).GetCell(20).SetCellValue("-");
                    #endregion

                    #region 2.農產品&畜產
                    comm.CommandText = @"
                        SELECT
	                        config.id as configid,
	                        config.name as name, 
	                        config.rownum as rownum,
	                        config.type as type, 
	                        config.isntvalid as isntvalid,
	                        config.iskgvalid as iskgvalid,
	                        config.source as source,
                            CASE WHEN display.configId IS NULL
                                THEN 'N'
                                ELSE 'Y'
                            END as ischecked
                        FROM
	                        config
                        LEFT JOIN(
	                        SELECT configid
	                        FROM display_checked
	                        INNER JOIN display
	                        ON display_checked.displayid = display.id
	                        AND display.name = 'new'
                        )display
                        ON config.id = display.configId
                        WHERE  
	                        isTrack = 'Y'
                    ";
                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@displayname", displayname);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Clear();
                        sda.Fill(dt);
                    }
                    List<Product> productList = new List<Product>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        Product product = new Product() {
                            id = (int)dr["configid"],
                            name = dr["name"].ToString(),
                            rownum = (int)dr["rownum"] - 1,
                            type = dr["type"].ToString(),
                            isntvalid = (dr["isntvalid"].ToString() == "Y"),
                            iskgvalid = (dr["iskgvalid"].ToString() == "Y"),
                            source = dr["source"].ToString(),
                            ischecked = (dr["ischecked"].ToString() == "Y")
                        };

                        productList.Add(product);
                    }

                    foreach (Product product in productList)
                    {
                        comm.Parameters.Clear();
                        StringBuilder sb = new StringBuilder();

                        #region (1)取得近兩周金額及交易量
                        switch (product.type)
                        {
                            //農產品
                            case ("FarmProduct"):
                            case ("FlowerProduct"):
                                sb.Append(@"
                                    SELECT
                                        config.name as productname,
                                        crops_price.year + '.' + crops_price.month + '.' + crops_price.days as times,
                                        crops_price.avg,
                                        crops_price.nt
                                    FROM
                                        config,
                                        crops,
                                        crops_price
                                    WHERE
                                        config.id = crops.configId
                                        AND crops.id = crops_price.cropId                                
                                        AND config.id = @configid
                                        AND crops_price.year + '.' + crops_price.month + '.' + crops_price.days in ( 
                                ");
                                break;
                            //畜產
                            case ("LivestockProduct"):
                                sb.Append(@"
                                    SELECT
                                        config.name as productname,
                                        livestocks_price.year + '.' + livestocks_price.month + '.' + livestocks_price.days as times,
                                        livestocks_price.avg,
                                        livestocks_price.nt,
                                        livestocks_price.kg
                                    FROM
                                        config,
                                        livestocks,
                                        livestocks_price
                                    WHERE
                                        livestocks.configid = config.id
                                        AND livestocks.id = livestocks_price.livestockid                                
                                        AND config.id = @configid
                                        AND livestocks.isTrack = 'Y'
                                        AND livestocks_price.year + '.' + livestocks_price.month + '.' + livestocks_price.days in ( 
                                ");
                                break;
                        }

                        for (var i = 0; i < 14; i++)
                        {
                            sb.Append("@date" + i.ToString() + ",");
                            comm.Parameters.AddWithValue("@date" + i.ToString(), basedate.AddDays(i * -1).ToString("yyyy.M.d", tc));
                        }
                        sb.Length--;
                        sb.Append(")");
                        comm.Parameters.AddWithValue("@configid", product.id);
                        comm.CommandText = sb.ToString();
                        using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                        {
                            dt.Clear();
                            sda.Fill(dt);
                        }
                        
                        foreach (DataRow dr in dt.Rows)
                        { 
                            int daydiff = (int)Math.Floor((basedate - DateTime.ParseExact(dr["times"].ToString(), "yyyy.M.d", tc)).TotalDays);

                            if (product.isntvalid)
                            {
                                //毛豬&閹公羊
                                if(product.iskgvalid)
                                {
                                    product.tv_day[daydiff] += Double.Parse(dr["nt"].ToString());
                                    product.tkv_day[daydiff] += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["kg"].ToString());
                                    product.tov_day[daydiff] += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["avg"].ToString()) * Double.Parse(dr["kg"].ToString());
                                }
                                else
                                {
                                    product.tv_day[daydiff] += Double.Parse(dr["nt"].ToString());
                                    product.tov_day[daydiff] += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["avg"].ToString());
                                }

                            }
                            //若產品不記錄產量則每天交易量以1計算
                            else
                            {                    
                                product.tv_day[daydiff] += 1;
                                product.tov_day[daydiff] += 1 * Double.Parse(dr["avg"].ToString());
                            }

                        }

                        for (int i = 0; i < 14; i++)
                        {
                            //本週
                            if (i < 7)
                            {
                                if (product.tv_day[i] > 0) {
                                    product.td_this++;
                                }
                                //有交易量&有重量
                                if (HasValue(product.tov_day[i] / product.tkv_day[i])) {
                                    product.tov_this += product.tov_day[i];
                                    product.tkv_this += product.tkv_day[i];
                                    product.tv_this += product.tv_day[i];
                                    sheet.GetRow(product.rownum).GetCell(18 - i).SetCellValue(product.tov_day[i] / product.tkv_day[i]);
                                    
                                }
                                //有交易量&無重量
                                else if (HasValue(product.tov_day[i] / product.tv_day[i]))
                                {
                                    product.tov_this += product.tov_day[i];
                                    product.tv_this += product.tv_day[i];
                                    sheet.GetRow(product.rownum).GetCell(18 - i).SetCellValue(product.tov_day[i] / product.tv_day[i]);
                                }
                                //無交易量&無重量
                                else
                                {
                                    //努比亞雜交閹公羊取上一個有價格的日期並印出日期(M/d)
                                    if (product.name == "努比亞雜交閹公羊")
                                    {
                                        var diff = 0;
                                        for (int j = 13; j >= i; j--)
                                        {
                                            if (product.tov_day[j] != 0)
                                            {
                                                diff = j;
                                            }
                                        }
                                        sheet.GetRow(product.rownum).GetCell(18 - i).SetCellValue(basedate.AddDays(diff * -1).ToString("(M/d)"));
                                        //如果是本週第一天，取上一個有價格的日期並印出價格
                                        if (i == 6) {
                                            sheet.GetRow(product.rownum).GetCell(18 - i).SetCellValue(product.tov_day[diff] / product.tkv_day[diff]);
                                        }
                                    }
                                    else {
                                        sheet.GetRow(product.rownum).GetCell(18 - i).SetCellValue("-");
                                    }                                    
                                }
                                //日量
                                if (product.tv_day[i] > 1)
                                {
                                    sheet.GetRow(product.rownum).GetCell(29 - i).SetCellValue(product.tv_day[i]);
                                }
                                else
                                {
                                    sheet.GetRow(product.rownum).GetCell(29 - i).SetCellValue("-");
                                }
                            }
                            //上週
                            else
                            {
                                if (product.tv_day[i] > 0)
                                {
                                    product.td_last++;
                                }
                                if (HasValue(product.tov_day[i] / product.tkv_day[i]))
                                {
                                    product.tov_last += product.tov_day[i];
                                    product.tkv_last += product.tkv_day[i];
                                    product.tv_last += product.tv_day[i];
                                }
                                else if(HasValue(product.tov_day[i] / product.tv_day[i]))
                                {
                                    product.tov_last += product.tov_day[i];
                                    product.tv_last += product.tv_day[i];
                                }
                            }
                        }
                        #endregion

                        #region (2)取得去年同月平均價格
                        switch (product.type) {
                            //農產品
                            case ("FarmProduct"):
                            case ("FlowerProduct"):
                                comm.CommandText = @"
									SELECT
                                        crops_price.nt AS nt,
                                        crops_price.avg AS avg                                       
                                    FROM
                                        config,
                                        crops,
                                        crops_price
                                    WHERE
                                        config.id = crops.configId
                                        AND crops.id = crops_price.cropId                                
                                        AND config.id = @configid
                                        AND crops_price.year = @year
                                        AND crops_price.month = @month
                                ";

             //                   if (!product.isntvalid && product.source == "產地價格")
             //                   {
             //                       comm.CommandText = @"
									    //SELECT
             //                               0 AS nt,
             //                               monthly_local.avg AS avg                                       
             //                           FROM                                        
             //                               monthly_local
             //                           LEFT JOIN 
										   // config   
									    //ON 
             //                               monthly_local.configid = config.id  
             //                           WHERE                         
             //                               config.id = @configid
             //                               AND monthly_local.year = @year
             //                               AND monthly_local.month = @month
             //                       ";
             //                   }
                                break;
                            //畜產
                            case ("LivestockProduct"):
                                comm.CommandText = @"
                                    SELECT
                                        livestocks_price.nt AS nt,
                                        livestocks_price.avg AS avg,
                                        livestocks_price.kg AS kg                  
                                    FROM
                                        config,
                                        livestocks,
                                        livestocks_price
                                    WHERE
                                        config.id = livestocks.configId
                                        AND livestocks.id = livestocks_price.livestockId                                
                                        AND config.id = @configid
                                        AND livestocks_price.year = @year
                                        AND livestocks_price.month = @month
                                        AND livestocks.isTrack = 'Y'
                                ";
                                break;
                        }

                        comm.Parameters.Clear();
                        comm.Parameters.AddWithValue("@year", basedate.AddYears(-1).ToString("yyyy", tc));
                        comm.Parameters.AddWithValue("@month", basedate.AddYears(-1).Month);
                        comm.Parameters.AddWithValue("@configid", product.id);

                        using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                        {
                            dt.Clear();
                            sda.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt.Rows) {

                                    if (product.isntvalid)
                                    {
                                        //毛豬&閹公羊
                                        if (product.iskgvalid)
                                        {
                                            product.tv_same_month_last_year += Double.Parse(dr["nt"].ToString());
                                            product.tkv_same_month_last_year += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["kg"].ToString());
                                            product.tov_same_month_last_year += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["kg"].ToString()) * Double.Parse(dr["avg"].ToString());
                                        }
                                        else {
                                            product.tv_same_month_last_year += Double.Parse(dr["nt"].ToString());
                                            product.tov_same_month_last_year += Double.Parse(dr["nt"].ToString()) * Double.Parse(dr["avg"].ToString());
                                        }
                                    }
                                    else {

                                        product.tv_same_month_last_year += 1;
                                        product.tov_same_month_last_year += Double.Parse(dr["avg"].ToString()) * 1;

                                    }
                                }
                            }
                        }
                        #endregion
                        //交易量
                        if (product.isntvalid)
                        {
                            sheet.GetRow(product.rownum).GetCell(19).SetCellValue(product.tv_this / product.td_this);
                            sheet.GetRow(product.rownum).GetCell(20).SetCellValue(100 * (product.tv_this / product.td_this - product.tv_last / product.td_last) / (product.tv_last / product.td_last));
                        }
                        else
                        {
                            sheet.GetRow(product.rownum).GetCell(19).SetCellValue("…");
                            sheet.GetRow(product.rownum).GetCell(20).SetCellValue("-");
                        }
                        //毛豬&閹公羊
                        if (product.iskgvalid && product.isntvalid)
                        {
                            //去年同月平均
                            sheet.GetRow(product.rownum).GetCell(6).SetCellValue(product.tov_same_month_last_year / product.tkv_same_month_last_year);
                            //這週平均
                            sheet.GetRow(product.rownum).GetCell(7).SetCellValue(product.tov_this / product.tkv_this);
                            //上週平均
                            sheet.GetRow(product.rownum).GetCell(22).SetCellValue(product.tov_last / product.tkv_last);
                            //這週與上周比較平均
                            sheet.GetRow(product.rownum).GetCell(11).SetCellValue(100 * ((product.tov_this / product.tkv_this) - (product.tov_last / product.tkv_last)) / (product.tov_last / product.tkv_last));
                        }
                        else {
                            //去年同月平均
                            sheet.GetRow(product.rownum).GetCell(6).SetCellValue(product.tov_same_month_last_year / product.tv_same_month_last_year);
                            //這週平均
                            sheet.GetRow(product.rownum).GetCell(7).SetCellValue(product.tov_this / product.tv_this);
                            //上週平均
                            sheet.GetRow(product.rownum).GetCell(22).SetCellValue(product.tov_last / product.tv_last);
                            //這週與上周比較平均
                            sheet.GetRow(product.rownum).GetCell(11).SetCellValue(100 * ((product.tov_this / product.tv_this) - (product.tov_last / product.tv_last)) / (product.tov_last / product.tv_last));
                        }
                    }
                    #endregion                   

                    #region 隱藏非產季欄位

                    for (var i = 11; i <= 126; i++) {

                        bool isHide = true;

                        if (productList.Any(p => p.ischecked == true && p.rownum == i)) {
                            isHide = false;
                        }
                        sheet.GetRow(i).ZeroHeight = isHide;
                    }
                    #endregion
                }
            }

            #region 設定excel檔名並寫入
            DateTime startDate = basedate.AddDays(-6);
            DateTime endDate = basedate;
            string dayofweek = "";
            //「價格週X」, X加一天為日報表的命名習慣
            switch (endDate.AddDays(1).DayOfWeek.ToString("d"))
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

            string file = startDate.ToString("yyyy.MM.dd",tc) + "-" + endDate.ToString("yyyy.MM.dd",tc) + "價格" + dayofweek + ".xlsx";
            if (!String.IsNullOrEmpty(path)) {
                file = Path.Combine(path , Path.GetFileName(file));
            }
            using (FileStream fs = File.Open(file, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(fs);
                fs.Close();
                Console.WriteLine("檔案路徑：" + file);                                
            }
            #endregion
        }

        private static bool HasValue(double value)
        {
            return !Double.IsNaN(value) && !Double.IsInfinity(value);
        }

    }

    class Product
    {
        public int id { get; set; }
        public string name { get; set; }
        public int rownum { get; set; }//在excel中的row編號
        public string type { get; set; }//產品種類
        public bool isntvalid { get; set; } //是否有交易量
        public bool iskgvalid { get; set; } //是否有重量
        public string source { get; set; } //產地價格或批發價格
        public bool ischecked { get; set; } //是否隱藏

        public int td_this = 0; //本周交易天數
        public int td_last = 0; //上周交易天數

        public double tv_this { get; set; }//本週總交易量
        public double tkv_this { get; set; }//本週總重量
        public double tov_this { get; set; }//本週總成交金額

        public double tv_last { get; set; }//上週總交易量
        public double tkv_last { get; set; }//上週總重量
        public double tov_last { get; set; }//上週總成交金額

        public double[] tv_day = new double[14];//當日交易量
        public double[] tkv_day = new double[14];//當日總重量
        public double[] tov_day = new double[14];//當日平均金額

        public double tv_same_month_last_year { get; set; }//去年同月總交易量
        public double tkv_same_month_last_year { get; set; }//去年同月總重量
        public double tov_same_month_last_year { get; set; }//去年同月總成交金額
    }
}
