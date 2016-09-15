using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace csv2dg
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable("tab0");
            int i=0;
            //for(int i=0;i<21;i++)
            //{
                //string a = i.ToString();
            DataColumn  a0 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a1 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a2 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a3 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a4 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a5 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a6 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a7 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a8 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a9 = new DataColumn(i++.ToString(), typeof(String));
            DataColumn a10 = new DataColumn(i++.ToString(), typeof(String));
            //}
            dt.Columns.AddRange(new DataColumn[] {a0,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10});
            string path = @"G:\project\excell_rozklad\rozklad.csv";
            string[] tab0 = File.ReadAllLines(path, Encoding.Default);
            string[] tab0Values = null;
            DataRow dr = null;
            for (i = 0; i < tab0.Length; i++)
            {
                if (!String.IsNullOrEmpty(tab0[i]))
                {
                    tab0Values = tab0[i].Split(';');
                    //создаём новую строку
                    dr = dt.NewRow();
                    for(int j=0;j<8;j++)
                    {string valp= tab0Values[j];
                        dr[j] = Regex.Replace(valp, " {2,}", " ");}
                    dt.Rows.Add(dr);
                }
            }

            dataGridView1.DataSource = dt;
  
                }

    }
}
