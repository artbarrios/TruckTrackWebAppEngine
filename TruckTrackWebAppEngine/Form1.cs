using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TruckTrackWebAppEngine
{
    public partial class Form1 : Form
    {
        public Form1(string appName)
        {
            InitializeComponent();
            this.Text = appName + " App Engine";
            AppCommon.StartAppEngine(richTextBox1);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            // set the richTextBox1 to autoscroll 
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }
    }
}

