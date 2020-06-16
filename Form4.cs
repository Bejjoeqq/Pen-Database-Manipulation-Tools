using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pen
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            this.TransparencyKey = (BackColor);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.Opacity > 0)
            {
                this.Opacity -= 0.01;
            }
            else
            {
                timer1.Enabled = false;
                this.Close();
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {
        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
