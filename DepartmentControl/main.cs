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
    public partial class main : Form
    {
        int sx = 1, sy = 1;
        public main()
        {
            InitializeComponent();
            
            label1.Text = "Вы вошли как: " + login.A;
           
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = сотрудникиBindingSource;
            dataGridView1.DataSource = сотрудникиBindingSource;
          

        }
       

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = сотрудникиBindingSource;
            dataGridView1.DataSource = сотрудникиBindingSource;
            radioButton7.Checked = false;
            radioButton8.Checked = false;
            radioButton10.Checked = false;
            button1.Enabled = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = отпускBindingSource;
            dataGridView1.DataSource = отпускBindingSource;
            radioButton7.Checked = false;
            radioButton8.Checked = false;
            radioButton10.Checked = false;
            button1.Enabled = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = больничныйBindingSource;
            dataGridView1.DataSource = больничныйBindingSource;
            radioButton7.Checked = false;
            radioButton8.Checked = false;
            radioButton10.Checked = false;
            button1.Enabled = false;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = командировкаBindingSource;
            dataGridView1.DataSource = командировкаBindingSource;
            radioButton7.Checked = false;
            radioButton8.Checked = false;
            radioButton10.Checked = false;
            button1.Enabled = false;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = true;
            bindingNavigator1.BindingSource = штрафыBindingSource;
            dataGridView1.DataSource = штрафыBindingSource;
            radioButton7.Checked = false;
            radioButton8.Checked = false;
            radioButton10.Checked = false;
            button1.Enabled = false;
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = false;
            bindingNavigator1.BindingSource = selectWorkerhospitalsBindingSource;
            dataGridView1.DataSource = selectWorkerhospitalsBindingSource;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            button1.Enabled = true;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = false;
            bindingNavigator1.BindingSource = selectWorkerLateBindingSource;
            dataGridView1.DataSource = selectWorkerLateBindingSource;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            button1.Enabled = true;
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            bindingNavigator1.Enabled = false;
            bindingNavigator1.BindingSource = selecCountLateDateBindingSource;
            dataGridView1.DataSource = selecCountLateDateBindingSource;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            button1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int x = label1.Location.X; //текущее положение по Х
            int y = label1.Location.Y; //текущее положение по У
            int w = label1.Width; // расмеры лейбла что бы контролировать не выходит
            int h = label1.Height; //ли он за границы

            int maxx = panel1.Width; //размеры панели по которой бегает лейбл
            int maxy = panel1.Height;

            //проверка не выходит ли при следуэщем шаге лейбл за границы панели
            if (x + sx < 0|| x + sx + w > maxx) sx = -sx; else x += sx;
            if (y + sy < 0 || y + sy + h > maxy) sy = -sy; else y += sy;
            //помещает заново лейбл, но визуально его перемещает циклически с интервалом
            //заданого в таймере
            label1.Location = new Point(x, y);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }

                StartRow++;

                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void main_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.selecCountLateDate' table. You can move, or remove it, as needed.
            this.selecCountLateDateTableAdapter.Fill(this.отдел_кадровDataSet.selecCountLateDate);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.selectWorkerLate' table. You can move, or remove it, as needed.
            this.selectWorkerLateTableAdapter.Fill(this.отдел_кадровDataSet.selectWorkerLate);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.selectWorkerhospitals' table. You can move, or remove it, as needed.
            this.selectWorkerhospitalsTableAdapter.Fill(this.отдел_кадровDataSet.selectWorkerhospitals);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.Штрафы' table. You can move, or remove it, as needed.
            this.штрафыTableAdapter.Fill(this.отдел_кадровDataSet.Штрафы);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.Командировка' table. You can move, or remove it, as needed.
            this.командировкаTableAdapter.Fill(this.отдел_кадровDataSet.Командировка);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.Больничный' table. You can move, or remove it, as needed.
            this.больничныйTableAdapter.Fill(this.отдел_кадровDataSet.Больничный);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.Отпуск' table. You can move, or remove it, as needed.
            this.отпускTableAdapter.Fill(this.отдел_кадровDataSet.Отпуск);
            // TODO: This line of code loads data into the 'отдел_кадровDataSet.Сотрудники' table. You can move, or remove it, as needed.
            this.сотрудникиTableAdapter.Fill(this.отдел_кадровDataSet.Сотрудники);

        }
    }
}
