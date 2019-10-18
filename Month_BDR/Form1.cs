using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace Month_BDR
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Visible = true;
                label2.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                label3.Visible = true;
                label3.Text = openFileDialog2.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if ((radioButton1.Checked == false) && (radioButton2.Checked == false) && (radioButton3.Checked == false) && (radioButton4.Checked == false) && (radioButton5.Checked == false))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран филиал!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;

            }
            ///////////////////////////////

            if (comboBox1.Text.Length==0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран месяц!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ///////////////////////////////

            /////Создание объекта Фин. отчет по филиалу
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            /////////

            /////Создание объекта БДР
            Microsoft.Office.Interop.Excel.Application ObjExcel1 = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook1 = ObjExcel1.Workbooks.Open(openFileDialog2.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet1;
            ObjWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook1.Sheets["БДР СВОД"];
            /////////

            string code_bdr="";
            progressBar1.Maximum = ObjWorkSheet.UsedRange.Rows.Count;
            
            //Открытие файла Фин. отчет по филиалу и его чтение
            for (int i = 9; i <= ObjWorkSheet.UsedRange.Rows.Count; i++)
            {
                string article = Convert.ToString(ObjWorkSheet.Cells[i, 1].value);
                string[] num_code = article.Split(' ');//Разбивка строки на слова, чтбы вычленить первое слово, т.е. код
                
                string sum_pre = Convert.ToString(ObjWorkSheet.Cells[i, 5].value);//Чтение суммы
                string[] summ = sum_pre.Split(' ');//Разбивка строки на слова, чтбы вычленить первое слово, т.е. сумму и удалить из нее апостроф
                string sum=summ[0].Replace("'", "");//Удаление апострофа из суммы
                

                    
                for (int u = 11; u <= 611; u++)
                {
                    string code = Convert.ToString(ObjWorkSheet1.Cells[u, 2].value);
                    if (code != null) { code_bdr = code.Remove(0, 1); }//Удаление первого символа из прочитанного кода, т.е. точки
                    else continue;
                    
                    if (num_code[0] == code_bdr)
                    {
                        for (int m = 44; m <= 507; m++)
                        {
                            string month = Convert.ToString(ObjWorkSheet1.Cells[4, m].value);
                            if (month == comboBox1.Text)
                            {
                                for (int b = m; b <= m + 20; b++)
                                {
                                    string res_check="";
                                    string branch = Convert.ToString(ObjWorkSheet1.Cells[5, b].value);
                                    //Определение выделенного radiogroup для понимания какой филиал выбран
                                    for (int p = 0; p < 3; p++)
                                    {
                                        if (((RadioButton)groupBox1.Controls[p]).Checked == true) { res_check = ((RadioButton)groupBox1.Controls[p]).Text; }
                                    }
                                    ///////////////////////

                                    if (branch == res_check)
                                    {
                                        string temp_res = Convert.ToString(ObjWorkSheet1.Cells[u, b + 1].value);
                                        double ress = Double.Parse(sum, CultureInfo.InvariantCulture) + Convert.ToDouble(temp_res);
                                        ObjWorkSheet1.Cells[u, b + 1] = ress;
                                                                                
                                    }
                                }
                            }
                             
                        }
                        
                    }

                }
           
                progressBar1.Value = progressBar1.Value + 1;
                
            }

            ObjWorkBook.Close();
            ObjExcel.Quit();
            ObjWorkBook = null;
            ObjWorkSheet = null;
            ObjExcel = null;

            ObjExcel1.Visible = true;

            GC.Collect();
        }
    }
}
