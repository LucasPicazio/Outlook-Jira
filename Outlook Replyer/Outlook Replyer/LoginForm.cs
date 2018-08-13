using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outlook_Replyer
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Main f1 = new Main();
            Main.pwd = textBox2.Text;
            this.Visible = false;
            f1.ShowDialog();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = Environment.UserName;
        }
    }
}
