using System;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Media;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.ComponentModel;
using System.Globalization;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Media.Animation;
using System.Windows.Input;

namespace DailyReportGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        TextBoxOutputter outputter;
        DataTable dt = new DataTable();
        DataTable dtDataFrame = new DataTable();
        DataTable dtDataFrame2 = new DataTable();
        CultureInfo tc = new CultureInfo("zh-TW");

        bool ToSend = true;

        string checkRadioButton = string.Empty;

        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();
            RefleshInputGroup(DateTime.Now);
            RefleshCombo(string.Empty, string.Empty);
            RefleshListBox();
            RefleshDataFrame();
            Days.IsChecked = true;

            dtDataFrame.Columns[0].ColumnName = "日報表品項";
            dtDataFrame.Columns[1].ColumnName = "細項及作物代碼";
            dtDataFrame.Columns[2].ColumnName = "追蹤市場";
            dtDataFrame.Columns[3].ColumnName = "列號";

            dtDataFrame2.Columns.Add("產品名稱");            
            dtDataFrame2.Columns.Add("日期");
            dtDataFrame2.Columns.Add("平均價格");
            dtDataFrame2.Columns.Add("平均交易量");
            dtDataFrame2.Columns.Add("單位");


            DataFrame2.ItemsSource = dtDataFrame2.DefaultView;
            DataFrame2.CanUserAddRows = false;
            DataFrame2.CanUserResizeColumns = true;
            DataFrame2.CanUserResizeRows = false;
            DataFrame2.CanUserResizeColumns = false;

        }

        public List<string> GetComboItems(string isDefault)
        {
            var list = new List<string>();

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {              
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = "SELECT DISTINCT name AS name FROM display WHERE isdeleted = 'N' AND isDefault = @isDefault";
                    comm.Parameters.AddWithValue("@isDefault", isDefault);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    foreach (DataRow dr in dt.Rows) {
                        list.Add(dr["name"].ToString());
                    }
                }
            }
            return list;
        }

        public List<ConfigItem> GetListBoxItems() {

            var list = new List<ConfigItem>();

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                        SELECT
                            config.id as id,
                            config.name as name, 
                            config.type as type, 
                            config.isntvalid as isntvalid,
                            config.iskgvalid as iskgvalid
                        FROM
                            config
                        WHERE  
                            isTrack = 'Y'
                    ";

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    foreach (DataRow dr in dt.Rows)
                    {
                        list.Add(new ConfigItem()
                        {
                            Id = (int)dr["id"],
                            Name = (string)dr["name"],
                            Type = (string)dr["type"],
                            isNtValid = (string)dr["isntvalid"] == "Y",
                            isKgValid = (string)dr["iskgvalid"] == "Y"
                        });
                    }
                }
            }
            return list;
        }

        public void RefleshListBox() {

            var list = GetListBoxItems();

            listBox.ItemsSource = list;

        }

        public void RefleshCombo(string editableListName, string defalutListName) {

            var defalutList = GetComboItems("Y");
            var editableList = GetComboItems("N");

            comboBox.ItemsSource = editableList;
            comboBox2.ItemsSource = editableList;
            comboBoxDefalut.ItemsSource = defalutList;

            if (defalutList.Count > 0)
            {

                if (string.IsNullOrEmpty(defalutListName))
                {
                    comboBoxDefalut.SelectedIndex = 0;
                }
                else
                {
                    comboBoxDefalut.SelectedIndex = comboBoxDefalut.Items.IndexOf(defalutListName);
                }
            }

            if (editableList.Count > 0)
            {

                if (string.IsNullOrEmpty(editableListName))
                {
                    comboBox.SelectedIndex = 0;
                    comboBox2.SelectedIndex = 0;
                }
                else
                {
                    RefleshCheckBoxGroup(editableListName, defalutListName);
                    comboBox.SelectedIndex = comboBox.Items.IndexOf(editableListName);
                    comboBox2.SelectedIndex = comboBox.Items.IndexOf(editableListName);
                }
            }
            else {
                CheckBoxGroup.Children.Clear();
            }


        }

        public void RefleshInputGroup(DateTime date) {

            InputGroup.Children.Clear();

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                        SELECT A.name,A.type,A.cropid,A.livestockid,B.avg as avg_crops,C.avg as avg_livestocks FROM(
                            SELECT 
                                config.name as name,
                                config.type as type,
                                crops.id as cropid,
                                livestocks.id as livestockid 
                            FROM 
                                config
                            LEFT JOIN 
                                crops
                            ON config.id = crops.configId	
                            LEFT JOIN 
                                livestocks
                            ON config.id = livestocks.configId	
                            WHERE 
                                config.source= '產地價格'
                            AND 
                                config.isTrack = 'Y'
                        )A    
                        LEFT JOIN 
                        (
                            SELECT cropid, avg FROM crops_price
                            WHERE year = @year AND month = @month AND days = @days
                            GROUP BY cropid, avg
                        )B
                        ON A.cropId = B.cropId
                        LEFT JOIN 
                        (
                            SELECT livestockid, avg FROM livestocks_price
                            WHERE year = @year AND month = @month AND days = @days
                            GROUP BY livestockid, avg
                        )C
                        ON A.livestockid = C.livestockid
                    ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@year", date.AddDays(-1).ToString("yyyy", tc));
                    comm.Parameters.AddWithValue("@month", date.AddDays(-1).Month);
                    comm.Parameters.AddWithValue("@days", date.AddDays(-1).Day);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    foreach (DataRow dr in dt.Rows)
                    {
                        int cropid = Convert.IsDBNull(dr["cropid"]) ? -1 : (int)dr["cropid"];
                        int livestockid = Convert.IsDBNull(dr["livestockid"]) ? -1 : (int)dr["livestockid"];
                        string name = (string)dr["name"];
                        string type = (string)dr["type"];

                        decimal avg = -1;
                        if (!Convert.IsDBNull(dr["avg_crops"]))
                        {
                            avg = (decimal)dr["avg_crops"];
                        }
                        else if (!Convert.IsDBNull(dr["avg_livestocks"])) {

                            avg = (decimal)dr["avg_livestocks"];
                        }

                        GenerateInputsControls(cropid, livestockid, name, type, avg, InputGroup);
                    }

                }
            }

        }

        public void RefleshCheckBoxGroup(string editable_displayname, string defalut_displayname) {

            CheckBoxGroup.Children.Clear();

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                        DECLARE @TmpTable TABLE (
	                        id int, 
	                        name nvarchar(20),
	                        ischecked_editable char(1),
	                        ischecked_defalut char(1)
                        )
                        INSERT INTO @TmpTable(id, name, ischecked_editable, ischecked_defalut)
                        SELECT
	                        config.id,
	                        config.name,
	                        'N',
	                        'N'
                        FROM config                     
                        WHERE
	                        config.isTrack = 'Y'
	
                        UPDATE @TmpTable
                        SET ischecked_editable = 'Y'
                        WHERE id IN(
	                        SELECT configid 
	                        FROM display_checked
	                        LEFT JOIN display
	                        ON display.id = display_checked.displayid
	                        WHERE display.name = @editable_displayname
                        )
                        
                        UPDATE @TmpTable
                        SET ischecked_defalut = 'Y'
                        WHERE id IN(
	                        SELECT configid 
	                        FROM display_checked
	                        LEFT JOIN display
	                        ON display.id = display_checked.displayid
	                        WHERE display.name = @defalut_displayname
                        )
                        SELECT * FROM @TmpTable                      
                    ";

                    comm.Parameters.AddWithValue("@editable_displayname", editable_displayname);
                    comm.Parameters.AddWithValue("@defalut_displayname", defalut_displayname);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    foreach (DataRow dr in dt.Rows)
                    {
                        int configid = (int)dr["id"];
                        string name = (string)dr["name"];
                        bool ischecked = (string)dr["ischecked_editable"] == "Y" ? true : false;
                        bool ismarked = (string)dr["ischecked_defalut"] == "Y" ? true : false;

                        GenerateCheckBoxControls(configid, name, ischecked, ismarked);
                    }

                }
            }

        }

        public void RefleshDataFrame()
        {

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                        SELECT * FROM(
	                        SELECT A.name AS n,
		                        STUFF(
			                        (SELECT DISTINCT ',' + BB.name + '(' + BB.code + ')'
			                        FROM config AA, crops BB
			                        WHERE AA.id = BB.configId
			                        AND AA.id = A.id
			                        FOR XML PATH('')), 1, 1, '') AS n2 ,
		                        ISNULL(STUFF(
			                        (SELECT DISTINCT ',' + C.name
			                        FROM market C, config AA, crops BB
			                        WHERE AA.id = BB.configId
			                        AND BB.marketId = C.id
			                        AND AA.id = A.id
			                        FOR XML PATH('')), 1, 1, ''),'全部市場') AS m,
		                        A.rowNum As r      
	                        FROM config A
	                        LEFT JOIN crops B
	                        ON A.id = B.configid
	                        WHERE B.name IS NOT NULL
	                        AND B.code IS NOT NULL
                        )S
                        GROUP BY n,n2,m,r
                        ORDER BY r                 
                    ";

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dtDataFrame.Clear();
                        sda.Fill(dtDataFrame);
                        DataFrame.ItemsSource = dtDataFrame.DefaultView;
                        DataFrame.Loaded += SetMinWidths;
                        DataFrame.LoadingRow += SetRowColor;
                        DataFrame.CanUserAddRows = false;
                        DataFrame.CanUserResizeColumns = false;
                        DataFrame.CanUserResizeRows = false;
                    }                   
                }
            }
        }

        public void SetMinWidths(object source, EventArgs e)
        {
            foreach (var column in DataFrame.Columns)
            {
                column.MinWidth = column.ActualWidth;
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            }
        }

        private void SetRowColor(object sender, DataGridRowEventArgs e)
        {
            int index = DataFrame.ItemContainerGenerator.IndexFromContainer(e.Row);
            if (index % 2 == 0)
            {
                e.Row.Background = new SolidColorBrush(Colors.LightGray);
            }
            else
            {
                e.Row.Background = new SolidColorBrush(Colors.WhiteSmoke);
            }
        }

        public void GenerateInputsControls(int cropid, int livestockid, string name, string type,decimal avg, WrapPanel group)
        {          
            Thickness tn = new Thickness();
            tn.Top = 5;
            tn.Bottom = 5;
            tn.Left = 5;
            tn.Right = 5;

            System.Windows.Controls.Label lb = new System.Windows.Controls.Label();
            lb.Width = 50;
            lb.Content = name;
            lb.Margin = tn;
            lb.Height = 24;

            System.Windows.Controls.TextBox tb = new System.Windows.Controls.TextBox();
            tb.DataContext = new 
            {
                name = name,
                cropid = cropid,
                livestockid = livestockid,
                type = type,
                avg = avg

            };
            tb.Width = 60;
            tb.Margin = tn;
            tb.Height = 24;
            switch (type) {
                case "FarmProduct":
                default:
                    tb.Background = Brushes.Aquamarine;
                    break;
                case "LivestockProduct":
                    tb.Background = Brushes.PaleTurquoise;
                    break;
            }

            group.Children.Add(lb);
            group.Children.Add(tb);

            group.UpdateLayout();
        }

        public void GenerateCheckBoxControls(int configid, string name, bool ischecked, bool ismarked){

            Thickness tn = new Thickness();
            tn.Top = 5;
            tn.Bottom = 5;
            tn.Left = 5;
            tn.Right = 5;

            Thickness tn2 = new Thickness();
            tn2.Top = 12;
            tn2.Bottom = 0;
            tn2.Left = 5;
            tn2.Right = 5;

            System.Windows.Controls.StackPanel sp = new System.Windows.Controls.StackPanel();
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal;

            System.Windows.Controls.Label lb = new System.Windows.Controls.Label();
            lb.Width = 70;
            lb.Content = name;
            lb.Margin = tn;
            lb.FontSize = 12;
            if (ismarked) {
                lb.Background = new SolidColorBrush(Colors.LightBlue);
            }            

            System.Windows.Controls.CheckBox cb = new System.Windows.Controls.CheckBox();
            cb.DataContext = new
            {
                configid = configid
            };

            cb.Margin = tn2;
            cb.IsChecked = ischecked;

            sp.Children.Add(cb);
            sp.Children.Add(lb);
            CheckBoxGroup.Children.Add(sp);

            CheckBoxGroup.UpdateLayout();
        }

        public MainWindow()
        {
            InitializeComponent();
            tc.DateTimeFormat.Calendar = new TaiwanCalendar();
            ReportDate.DisplayDate = DateTime.Now.Date;
            ReportDate.DisplayDateEnd = DateTime.Now.Date;
            ImportDate.DisplayDate = DateTime.Now.Date;
            ImportDate.DisplayDateEnd = DateTime.Now.Date;
            StartDate.DisplayDate = DateTime.Now.Date;
            StartDate.DisplayDateEnd = DateTime.Now.Date;
            EndDate.DisplayDate = DateTime.Now.Date;
            EndDate.DisplayDateEnd = DateTime.Now.Date;
            DataContext = this;         
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Tab1.IsSelected) {
                outputter = new TextBoxOutputter(LogBox);
                Console.SetOut(outputter);
            }
            if (Tab2.IsSelected)
            {
                outputter = new TextBoxOutputter(LogBox2);
                Console.SetOut(outputter);
            }
        }

        public void Generate(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(PathText.Text)) {
                Console.WriteLine("請選擇產生路徑！");
                return;
            }

            if (String.IsNullOrEmpty(comboBox2.SelectedItem.ToString())) {
                Console.WriteLine("請新增顯示清單！");
                return;
            }

            LogBox.Text = String.Empty;
            var selectedDate = Convert.ToDateTime(ReportDate.SelectedDate);

            Exe getFarmProduct = new Exe()
            {
                Name = "GetFarmProduct.exe",
                Arguments = selectedDate.AddDays(-14).ToString("yyyy/MM/dd") + " " + selectedDate.ToString("yyyy/MM/dd"),
                Slog = "正在抓取農產品資訊......",
                Elog = "完成抓取農產品資訊!"
            };
            Exe dailyReport = new Exe()
            {
                Name = "DailyReport.exe",
                Arguments = selectedDate.ToString("yyyy/MM/dd") + " " + PathText.Text + " " + comboBox2.SelectedItem.ToString(),
                Slog = "正在產生日報表......",
                Elog = "完成!"
            };
            Exe getRicePrice = new Exe()
            {
                Name = "GetRicePrice.exe",
                Arguments = selectedDate.AddMonths(-1).ToString("yyyy/MM/dd") + " " + selectedDate.ToString("yyyy/MM/dd"),
                Slog = "正在抓取糧價資訊......",
                Elog = "完成抓取糧價資訊!"
            };

            Exe getLiveStocksProduct = new Exe()
            {
                Name = "GetLiveStocksProduct.exe",
                Arguments = selectedDate.AddDays(-14).ToString("yyyy/MM/dd") + " " + selectedDate.ToString("yyyy/MM/dd"),
                Slog = "正在抓取畜產資訊......",
                Elog = "完成抓取畜產資訊!"
            };
            Exe getFlowerProduct = new Exe()
            {
                Name = "GetFlowerProduct.exe",
                Arguments = selectedDate.AddDays(-14).ToString("yyyy/MM/dd") + " " + selectedDate.ToString("yyyy/MM/dd"),
                Slog = "正在抓取花卉資訊......",
                Elog = "完成抓取花卉資訊!"
            };

            //Exe getLocalPriceMonthly = new Exe()
            //{
            //    Name = "GetLocalPriceMonthly.exe",
            //    Arguments = null,
            //    Slog = "正在抓取產地價格月資料......",
            //    Elog = "完成抓取產地價格月資料!"
            //};

            Stopwatch watch = new Stopwatch();

            Thread t6 = new Thread(() => { LaunchCommandLineApp(dailyReport); watch.Stop(); Console.WriteLine("執行時間：" + watch.Elapsed.ToString(@"mm\:ss")); });
            //Thread t5 = new Thread(() => { LaunchCommandLineApp(getLocalPriceMonthly); t6.Start(); });
            Thread t4 = new Thread(() => { LaunchCommandLineApp(getLiveStocksProduct); t6.Start(); });
            Thread t3 = new Thread(() => { LaunchCommandLineApp(getFlowerProduct); t4.Start(); });
            Thread t2 = new Thread(() => { LaunchCommandLineApp(getFarmProduct); t3.Start(); });
            Thread t1 = new Thread(() => { LaunchCommandLineApp(getRicePrice); t2.Start(); });            
            watch.Start();
            t1.Start();
        }

        private void LaunchCommandLineApp(Exe exe)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = System.IO.Path.Combine(Environment.CurrentDirectory, "exe");
            startInfo.FileName = exe.Name;
            startInfo.Arguments = exe.Arguments;

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    Console.Write(exe.Slog);
                    exeProcess.WaitForExit();
                    Console.WriteLine(exe.Elog);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(startInfo.StandardOutputEncoding);
                Console.WriteLine(e);
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var picker = sender as DatePicker;            
            if (picker.Name == "ImportDate" && IsInitialized)
            {
                Console.WriteLine("你選擇了匯入日期：" + Convert.ToDateTime(picker.SelectedDate).ToString("yyyy/MM/dd"));
                RefleshInputGroup(Convert.ToDateTime(picker.SelectedDate));
            }
            if (picker.Name == "ReportDate")
            {
                Console.WriteLine(String.Format("你選擇了產生日期：{0} - {1}" ,Convert.ToDateTime(picker.SelectedDate).AddDays(-6).ToString("yyyy/MM/dd"), Convert.ToDateTime(picker.SelectedDate).ToString("yyyy/MM/dd")));
            }

        }

        private void OpenPath_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (!String.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    PathText.Text = fbd.SelectedPath;
                }
            }
        }               

        public void Import(object sender, RoutedEventArgs e)
        {
            LogBox2.Text = String.Empty;

            decimal value;

            bool isChecked = true;

            StringBuilder sb = new StringBuilder();

            foreach (System.Windows.Controls.TextBox tb in FindVisualChildren<System.Windows.Controls.TextBox>(InputGroup))
            {
                decimal avg = (decimal)TypeDescriptor.GetProperties(tb.DataContext)["avg"].GetValue(tb.DataContext);
                string name = (string)TypeDescriptor.GetProperties(tb.DataContext)["name"].GetValue(tb.DataContext);

                if (String.IsNullOrWhiteSpace(tb.Text)) {
                    continue;
                }
                else if (Decimal.TryParse(tb.Text, out value))
                {
                    if (value != avg) {
                        sb.Append(String.Format("{0}：{1} → {2}\n", name, avg == -1? "無價格": avg.ToString(), value.ToString()));
                    }                   
                }
                else
                {
                    Console.WriteLine(name + "輸入有誤!請重新輸入。");
                    tb.Text = String.Empty;
                    isChecked = false;
                }
            }

            if (isChecked)
            {

                MessageBoxResult result = System.Windows.MessageBox.Show("以下為比較前一天價格之變動項目，確定要匯入資料？\n"+ sb, "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
                {

                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.Connection = conn;
                        conn.Open();

                        foreach (System.Windows.Controls.TextBox tb in FindVisualChildren<System.Windows.Controls.TextBox>(InputGroup))
                        {
                            if (String.IsNullOrEmpty(tb.Text))
                            {
                                continue;
                            }
                            int cropid = (int)TypeDescriptor.GetProperties(tb.DataContext)["cropid"].GetValue(tb.DataContext);
                            int livestockid = (int)TypeDescriptor.GetProperties(tb.DataContext)["livestockid"].GetValue(tb.DataContext);
                            string type = (string)TypeDescriptor.GetProperties(tb.DataContext)["type"].GetValue(tb.DataContext);
                            string name = (string)TypeDescriptor.GetProperties(tb.DataContext)["name"].GetValue(tb.DataContext);

                            var date = Convert.ToDateTime(ImportDate.SelectedDate);

                            switch (type)
                            {
                                case "FarmProduct":
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
                                            VALUES(@cropid,@year,@month,@days,@avg,0,GETDATE())
                                        END
                                        ELSE
                                        UPDATE crops_price
                                        SET 
                                            avg = @avg,nt = 0,updateTime = GETDATE()
                                        WHERE cropid = @cropid
                                        AND year = @year
                                        AND month = @month
                                        AND days = @days      
                                    ";

                                    comm.Parameters.Clear();
                                    comm.Parameters.AddWithValue("@cropid", cropid);
                                    comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                                    comm.Parameters.AddWithValue("@month", date.Month);
                                    comm.Parameters.AddWithValue("@days", date.Day);
                                    comm.Parameters.AddWithValue("@avg", Convert.ToDecimal(tb.Text));
                                    comm.ExecuteNonQuery();

                                    break;
                                case "LivestockProduct":
                                    comm.CommandText = @"
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
                                            VALUES(@livestockid,@year,@month,@days,@avg,0,0,GETDATE())
                                        END
                                        ELSE
                                        UPDATE livestocks_price
                                        SET 
                                            avg = @avg,nt = 0,kg=0,updateTime = GETDATE()
                                        WHERE livestockid = @livestockid
                                        AND year = @year
                                        AND month = @month
                                        AND days = @days
                                    ";

                                    comm.Parameters.Clear();
                                    comm.Parameters.AddWithValue("@livestockid", livestockid);
                                    comm.Parameters.AddWithValue("@year", date.ToString("yyyy", tc));
                                    comm.Parameters.AddWithValue("@month", date.Month);
                                    comm.Parameters.AddWithValue("@days", date.Day);
                                    comm.Parameters.AddWithValue("@avg", Convert.ToDecimal(tb.Text));
                                    comm.ExecuteNonQuery();
                                    break;
                            }
                            Console.WriteLine(String.Format("成功匯入一筆{0}資料。日期：{1}", name, date.ToString("yyyy.MM.dd", tc)));
                        }
                    }
                }
            }

       
        }

        public void UpdateCheckboxInput(object sender, RoutedEventArgs e)
        {
            if (comboBox.Items.Count == 0)
            {
                System.Windows.MessageBox.Show("請先新增清單!");
                return;
            }

            MessageBoxResult result = System.Windows.MessageBox.Show("確定要套用此更新？", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                var name = comboBox.SelectedItem.ToString();

                DeleteDisplayItems(name);

                foreach (System.Windows.Controls.CheckBox cb in FindVisualChildren<System.Windows.Controls.CheckBox>(CheckBoxGroup))
                {
                    int configid = (int)TypeDescriptor.GetProperties(cb.DataContext)["configid"].GetValue(cb.DataContext);

                    if (cb.IsChecked.Value)
                    {
                        InsertDisplayItem(name, configid);
                    }
                }

                System.Windows.MessageBox.Show("更新成功！");
            }          
        }

        public void ShowTextBoxText_Farm(object sender, RoutedEventArgs e) {

            foreach (System.Windows.Controls.TextBox tb in FindVisualChildren<System.Windows.Controls.TextBox>(InputGroup)) {

                var avg = (decimal)TypeDescriptor.GetProperties(tb.DataContext)["avg"].GetValue(tb.DataContext);
                var type = (string)TypeDescriptor.GetProperties(tb.DataContext)["type"].GetValue(tb.DataContext);
                if (avg != -1 && type == "FarmProduct")
                {
                    tb.Text = avg.ToString();
                }
            }

        }

        public void ShowTextBoxText_Livestock(object sender, RoutedEventArgs e)
        {

            foreach (System.Windows.Controls.TextBox tb in FindVisualChildren<System.Windows.Controls.TextBox>(InputGroup))
            {

                var avg = (decimal)TypeDescriptor.GetProperties(tb.DataContext)["avg"].GetValue(tb.DataContext);
                var type = (string)TypeDescriptor.GetProperties(tb.DataContext)["type"].GetValue(tb.DataContext);
                if (avg != -1 && type=="LivestockProduct")
                {
                    tb.Text = avg.ToString();
                }
            }

        }

        public void ClearTextBoxText(object sender, RoutedEventArgs e)
        {

            foreach (System.Windows.Controls.TextBox tb in FindVisualChildren<System.Windows.Controls.TextBox>(InputGroup))
            {
                tb.Text = String.Empty;
            }

        }

        public void ComboBoxChanged(object sender, SelectionChangedEventArgs e)
        {            
            string editableListName = comboBox.SelectedItem as string;
            string defalutListName = comboBoxDefalut.SelectedItem as string;
            RefleshCombo(editableListName, defalutListName);
        }

        public void NewCombo(object sender, RoutedEventArgs e) {

            InputDialog inputDialog = new InputDialog();

            var answer = string.Empty;

            if (inputDialog.ShowDialog() == true)
            {
                answer = inputDialog.Answer;
            }
            else {
                return;
            }

            if (!String.IsNullOrEmpty(answer) & !comboBox.Items.Contains(answer))
            {
                InsertDisplay(answer);
                RefleshCombo(answer, string.Empty);
            }
            else {
                System.Windows.MessageBox.Show("此清單名稱有誤或已被使用!");
            }             
        }

        public void DeleteCombo(object sender, RoutedEventArgs e) {

            if (comboBox.Items.Count == 0) {
                System.Windows.MessageBox.Show("請先新增清單!");
                return;
            }

            MessageBoxResult result = System.Windows.MessageBox.Show("確定要刪除此清單？", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes) {

                var name = comboBox.SelectedItem.ToString();
                DeleteDisplayItems(name);
                DeleteDisplay(name);
                RefleshCombo(string.Empty, string.Empty);

            }
        }

        public void DeleteDisplayItems(string name)
        {

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                            DELETE FROM display_checked WHERE displayid = (SELECT id FROM display WHERE display.name = @name AND isdeleted = 'N')
                        ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@name", name);
                    comm.ExecuteNonQuery();
                }
            }
        }

        public void InsertDisplayItem(string name, int configid) {

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                            INSERT INTO display_checked(displayid,configid,updatetime) 
                            VALUES(
                                (SELECT id FROM display WHERE display.name = @name AND isdeleted = 'N'),
                                @configid,
                                GETDATE()
                            )
                        ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@name", name);
                    comm.Parameters.AddWithValue("@configid", configid);
                    comm.ExecuteNonQuery();
                }
            }

        }

        public void DeleteDisplay(string name)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                            UPDATE display SET isdeleted = 'Y'
                            WHERE name = @name
                        ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@name", name);
                    comm.ExecuteNonQuery();
                }
            }
        }

        public void InsertDisplay(string name)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                            INSERT INTO display(name, isdeleted, isDefault, createTime)
                            VALUES(@name, 'N', 'N', GETDATE())
                        ";

                    comm.Parameters.Clear();
                    comm.Parameters.AddWithValue("@name", name);
                    comm.ExecuteNonQuery();
                }
            }
        }

        public void GetResult_Click(object sender, RoutedEventArgs e) {

            if (ToSend) {

                ToSend = false;
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;


                dtDataFrame2.Clear();

                var selectedItems = listBox.SelectedItems.Cast<ConfigItem>().ToList();

                foreach (ConfigItem config in selectedItems)
                {
                    switch (config.Type)
                    {
                        case "FarmProduct":
                            GetCrop(config, "公斤");
                            break;
                        case "FlowerProduct":
                            GetCrop(config, "把");
                            break;
                        case "LivestockProduct":
                            GetLivestock(config);
                            break;
                    }
                }

                DataFrame2.ItemsSource = dtDataFrame2.DefaultView;
                DataFrame2.Columns[0].Width = 60;
                DataFrame2.Columns[1].Width = 70;
                DataFrame2.Columns[2].Width = 60;
                DataFrame2.Columns[3].Width = 60;
                DataFrame2.Columns[4].Width = 30;
                DataFrame2.Items.Refresh();

                ToSend = true;
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;

            }


        }

        public void GetCrop(ConfigItem config, string unit)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                                        SELECT year AS year,
                                               month AS month, 
                                               days AS days, 
                                               avg AS avg, 
                                               nt AS nt,
                                               CASE  
                                               WHEN days <= 11 THEN '/上旬' 
                                               WHEN days <= 21 THEN '/中旬'
                                               WHEN days <= 31 THEN '/下旬'
                                               END AS tendays
                                        FROM crops_price 
                                        WHERE 
                                           (year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2))*1 <= @enddate
                                        AND 
                                           (year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2))*1 >= @startdate
                                        AND 
                                            cropid IN (SELECT id FROM crops WHERE configid = @configid)
                                        ORDER BY 
                                            year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2)
                                        ";

                    comm.Parameters.AddWithValue("@startdate", Convert.ToDateTime(StartDate.SelectedDate).ToString("yyyyMMdd", tc));
                    comm.Parameters.AddWithValue("@enddate", Convert.ToDateTime(EndDate.SelectedDate).ToString("yyyyMMdd", tc));
                    comm.Parameters.AddWithValue("@configid", config.Id);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    if (checkRadioButton == "Days") {

                        var grouped = from row in dt.AsEnumerable()
                                  group row by new { year = row.Field<string>("year"), month = row.Field<string>("month"), days = row.Field<string>("days") } into g
                                  select new
                                  {
                                      date = g.Key.year + '/' + g.Key.month + '/' + g.Key.days ,
                                      nt = g.Sum(p => Convert.ToDecimal(p["nt"])),
                                      avg = config.isNtValid ?
                                               g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"])) :
                                               g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                  };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = unit;                            
                            dtDataFrame2.Rows.Add(row);
                        }
                    }


                    if (checkRadioButton == "TenDays")
                    {

                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year"), month = row.Field<string>("month"), tendays = row.Field<string>("tendays") } into g
                                      select new
                                      {
                                          date = g.Key.year + '/' + g.Key.month + g.Key.tendays,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])),
                                          avg = config.isNtValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = unit;
                            dtDataFrame2.Rows.Add(row);
                        }
                    }

                    if (checkRadioButton == "Month") {

                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year"), month = row.Field<string>("month") } into g
                                      select new
                                      {
                                          date = g.Key.year + '/' + g.Key.month,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])),
                                          avg = config.isNtValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = unit;
                            dtDataFrame2.Rows.Add(row);
                        }

                    }

                    if (checkRadioButton == "Year")
                    {

                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year")} into g
                                      select new
                                      {
                                          date = g.Key.year,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])),
                                          avg = config.isNtValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {
                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = unit;
                            dtDataFrame2.Rows.Add(row);
                        }

                    }
                }
            }
        }

        public void GetLivestock(ConfigItem config) {

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["COA_ConnStr"].ToString()))
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    conn.Open();

                    comm.CommandText = @"
                                        SELECT year AS year,
                                               month AS month, 
                                               days AS days, 
                                               avg AS avg, 
                                               nt AS nt,
                                               kg AS kg,
                                               CASE  
                                               WHEN days <= 11 THEN '/上旬' 
                                               WHEN days <= 21 THEN '/中旬'
                                               WHEN days <= 31 THEN '/下旬'
                                               END AS tendays
                                        FROM livestocks_price 
                                        WHERE 
                                           (year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2))*1 <= @enddate
                                        AND 
                                           (year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2))*1 >= @startdate
                                        AND 
                                            livestockid IN (SELECT id FROM livestocks WHERE configid = @configid)
                                        ORDER BY 
                                            year + RIGHT('00'+ISNULL(month,''),2) + RIGHT('00'+ISNULL(days,''),2)
                                        ";

                    comm.Parameters.AddWithValue("@startdate", Convert.ToDateTime(StartDate.SelectedDate).ToString("yyyyMMdd", tc));
                    comm.Parameters.AddWithValue("@enddate", Convert.ToDateTime(EndDate.SelectedDate).ToString("yyyyMMdd", tc));
                    comm.Parameters.AddWithValue("@configid", config.Id);

                    using (SqlDataAdapter sda = new SqlDataAdapter(comm))
                    {
                        dt.Reset();
                        sda.Fill(dt);
                    }

                    if (checkRadioButton == "Days")
                    {
                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year"), month = row.Field<string>("month"), days = row.Field<string>("days") } into g
                                      select new
                                      {
                                          date = g.Key.year + '/' + g.Key.month + '/' + g.Key.days ,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])) / dt.Rows.Count,
                                          ntkg = g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])),
                                          avg = config.isNtValid && config.isKgValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = "公斤";
                            dtDataFrame2.Rows.Add(row);
                        }
                    }

                    if (checkRadioButton == "TenDays")
                    {
                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year"), month = row.Field<string>("month"), tendays = row.Field<string>("tendays") } into g
                                      select new
                                      {
                                          date = g.Key.year + '/' + g.Key.month + g.Key.tendays,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])) / dt.Rows.Count,
                                          ntkg = g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])),
                                          avg = config.isNtValid && config.isKgValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = "公斤";
                            dtDataFrame2.Rows.Add(row);
                        }
                    }

                    if (checkRadioButton == "Month")
                    {
                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year"), month = row.Field<string>("month") } into g
                                      select new
                                      {
                                          date = g.Key.year + '/' + g.Key.month,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])) / dt.Rows.Count,
                                          ntkg = g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])),
                                          avg = config.isNtValid && config.isKgValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = "公斤";
                            dtDataFrame2.Rows.Add(row);
                        }
                    }

                    if (checkRadioButton == "Year")
                    {
                        var grouped = from row in dt.AsEnumerable()
                                      group row by new { year = row.Field<string>("year")} into g
                                      select new
                                      {
                                          date = g.Key.year,
                                          nt = g.Sum(p => Convert.ToDecimal(p["nt"])) / dt.Rows.Count,
                                          ntkg = g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])),
                                          avg = config.isNtValid && config.isKgValid ?
                                                   g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"]) * Convert.ToDecimal(p["avg"])) / g.Sum(p => Convert.ToDecimal(p["nt"]) * Convert.ToDecimal(p["kg"])) :
                                                   g.Sum(p => Convert.ToDecimal(p["avg"])) / g.Count()
                                      };

                        foreach (var group in grouped)
                        {

                            var row = dtDataFrame2.NewRow();
                            row["產品名稱"] = config.Name;
                            row["日期"] = group.date;
                            row["平均價格"] = group.avg.ToString("0.##");
                            row["平均交易量"] = config.isNtValid ? group.nt.ToString("0.##") : "無";
                            row["單位"] = "公斤";
                            dtDataFrame2.Rows.Add(row);
                        }
                    }
                }
            }
        }

        public void radioButton_Checked(object sender, RoutedEventArgs e)
        {
            checkRadioButton = (string)(sender as System.Windows.Controls.RadioButton).Name;
        }

        public static System.Collections.Generic.IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }


    }

    public class Exe
    {
        public string Name { get; set; }
        public string Arguments { get; set; }
        public string Slog { get; set; }
        public string Elog { get; set; }
    }

    public class ConfigItem {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public bool isNtValid { get; set; }
        public bool isKgValid { get; set; }
    }
}
