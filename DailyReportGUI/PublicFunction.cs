using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;

namespace DailyReportGUI
{
    class DataFrame
    {
        public DataGrid Frame { get; set; }

        public void SetMinWidths(object source, EventArgs e)
        {
            foreach (var column in Frame.Columns)
            {
                column.MinWidth = column.ActualWidth;
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            }
        }

        private void SetRowColor(object sender, DataGridRowEventArgs e)
        {
            int index = Frame.ItemContainerGenerator.IndexFromContainer(e.Row);
            if (index % 2 == 0)
            {
                e.Row.Background = new SolidColorBrush(Colors.LightGray);
            }
            else
            {
                e.Row.Background = new SolidColorBrush(Colors.WhiteSmoke);
            }
        }

    }
}
