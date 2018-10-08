using System;
using System.Text;
using System.IO;

namespace DailyReportGUI
{
    class TextBoxOutputter : TextWriter
    {
        System.Windows.Controls.TextBox textBox = null;

        public TextBoxOutputter(System.Windows.Controls.TextBox output)
        {
            textBox = output;
        }

        public override void Write(char value)
        {
            base.Write(value);
            textBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                textBox.AppendText(value.ToString());
            }));
        }

        public override Encoding Encoding
        {
            get { return System.Text.Encoding.UTF8; }
        }
    }
}