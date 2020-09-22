using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DepartmentControl
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();

            adminTableAdapter1.Fill(отдел_кадровDataSet1.Admin);
        }
        public static string A { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {

      
            A = отдел_кадровDataSet1.Admin.Where(p => p.email == textBox1.Text && p.password == textBox2.Text).Select(s => s.email).FirstOrDefault()?.ToString();
           
            if (A != null)
            {
                main f = new main();
                f.Show();
            }
            else
            {
                MessageBox.Show("Error");
            }

           
        }
    }
}
