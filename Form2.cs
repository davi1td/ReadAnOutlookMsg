using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ReadAnOutlookMsg
{
    public partial class Form2 : Form
    {
        public Form2(string bodyText, string bodyRTF)
        {
            InitializeComponent();
            this.textBox1.Text = bodyText;
            this.richTextBox1.Rtf = bodyRTF;
        }
    }
}