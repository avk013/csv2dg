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
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Office = Microsoft.Office.Core;

namespace csv2dg
{
    public partial class Form1 : Form
    {
//        private int applicationHwnd;
  //      private object ExcelApplication;

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
            //открываем файл
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Text Files|*.csv";
            openFileDialog1.Title = "фал после обработки ВБА";
            openFileDialog1.FileName = "rozklad";
            //MessageBox.Show("файл с сайтами");
            string path;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                path=openFileDialog1.FileName;
            //иначе по умолчанию
            else path = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv";
            //else path = @"G:\project\excell_rozklad\rozklad.csv";
            //string path = @"rozklad.csv";
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
                    {string valp= tab0Values[j].ToUpper();
                        dr[j] = Regex.Replace(valp, " {2,}", " ");}
                    dt.Rows.Add(dr);
                }
            }

            dataGridView1.DataSource = dt;
            //пытаемся почистить
            string[] badwords = {"розклад", "деканфак", "занятьфак" };
            string[] wastewords = { "ДНІ", "ПАРИ", "понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота", "неділя", "П’ЯТНИЦЯ" };
            // удаляем строки с плохими словами
            for (int k = 0; k < badwords.Length; k++)
                for (int ii = 0; ii < dt.Columns.Count; ii++)
                for (int j = 0; j < dt.Rows.Count; j++)   
                {
                 if (dt.Rows[j][ii].ToString().Replace(" ", string.Empty).ToUpper().Contains(badwords[k].ToUpper()))
                        { //dt.Rows[j][ii] = "";                    
                        dt.Rows.RemoveAt(j);j--;}
                }
            // удаляем ненужные слова
            for (int k = 0; k < wastewords.Length; k++)
                for (int ii = 0; ii < 2; ii++)
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                    //int ii = 0;
                        if (dt.Rows[j][ii].ToString().Replace(" ", string.Empty).ToUpper().Contains(wastewords[k].ToUpper()))                            
                            dt.Rows[j][ii] = dt.Rows[j][ii].ToString().Replace(wastewords[k].ToUpper(),"");
                    
                }
            ////////////

            ////////////////////
            //перестраиваем таблицу под новый формат
            DataTable dt1 = new DataTable("tab1");

            i = 0;
            DataColumn gr = new DataColumn(i++.ToString(), typeof(String));
            DataColumn date = new DataColumn(i++.ToString(), typeof(String));
            DataColumn nomer = new DataColumn(i++.ToString(), typeof(String));
            DataColumn displ = new DataColumn(i++.ToString(), typeof(String));
            DataColumn prpd = new DataColumn(i++.ToString(), typeof(String));
            dt1.Columns.AddRange(new DataColumn[] { gr, date, nomer, displ, prpd });
            DataRow dr1 = null;

            string[] groups = { "11 ГРУПА", "12 ГРУПА", "13 ГРУПА", "14 ГРУПА", "15 ГРУПА",
                "22 ГРУПА", "23 ГРУПА", "24 ГРУПА", "25 ГРУПА",
                "31 ГРУПА", "32 ГРУПА", "33 ГРУПА", "34 ГРУПА", "35 ГРУПА",
                "41 ГРУПА", "42 ГРУПА", "43 ГРУПА", "44 ГРУПА", "45 ГРУПА",
                "51 ГРУПА", "52 ГРУПА", "53 ГРУПА", "54 ГРУПА", "55 ГРУПА",
                "61 ГРУПА", "62 ГРУПА", "63 ГРУПА", "64 ГРУПА", "65 ГРУПА",
                "71 ГРУПА", "72 ГРУПА", "73 ГРУПА", "74 ГРУПА", "75 ГРУПА",
                "17 ГРУПА", "18 ГРУПА", "19 ГРУПА", "16 ГРУПА", "211ГРУПА",
                "212ГРУПА", "221ГРУПА", "222ГРУПА", "311 ГРУПА",
                "312 ГРУПА", "321 ГРУПА", "322 ГРУПА", "511 ГРУПА",
                "512 ГРУПА" };
            string[] razdel_v = { "ДОЦ.","ВИКЛ."};
         //   string[] stolb1bad = { "ДНІ", "ПАРИ"};
            //ищем ячеки с группой
            int stroka=0;
            foreach (string group in groups)
                for (int ii = 0; ii < dt.Columns.Count; ii++)
                    for (int j = 0; j < dt.Rows.Count; j++)
                        if (group.ToUpper().Replace(" ", string.Empty) == dt.Rows[j][ii].ToString().ToUpper().Replace(" ", string.Empty))
                            // нашли столбец ii  и начальную строку j
                                 { int strk=j+2;
                            while(strk<dt.Rows.Count&&dt.Rows[strk][0].ToString()!="")
                            {
                                //     for (int jj = stroka; jj < dt.Rows.Count; jj++)
                                //     {
                                //  for (int k = 0; k < stolb1bad.Length; k++)
                                //        foreach(string badword in stolb1bad)
                                //if (!dt.Rows[jj][ii].ToString().ToUpper().Contains(badword.ToUpper()))

                                //       {// textBox1.Text +=group;
                                //       textBox1.Text += group+dt.Rows[jj][0].ToString().ToUpper()+ badword+Environment.NewLine;
                                //textBox1.Text += dt.Rows[strk][0].ToString()+"+"+ii.ToString() + "+" + j.ToString() + "=" + group+ "_"+dt.Rows[strk][ii].ToString()+ Environment.NewLine;// + dt.Rows[jj][0].ToString().ToUpper() ;


          
                                dr1 = dt1.NewRow();
                                
                                {
                                    dr1[0] = group.ToUpper().Replace("ГРУПА", string.Empty).Replace(" ", string.Empty);
                                    // адапитировать дату в мускл
                                    string g= dt.Rows[strk][0].ToString().Replace(" ", string.Empty).ToUpper();
                                    int f = g.IndexOf(".", 0);
                                    if (f>0)dr1[1] = "2016" + g.Substring(f, g.Length - f)+"."+g.Substring(0, f);// dat
                                    //dr1[1] = "2016-"+dt.Rows[strk][0].ToString().Replace(" ", string.Empty).ToUpper();// dat
                                    dr1[2] = dt.Rows[strk][1].ToString().Replace(" ", string.Empty).ToUpper();  //nom
                                    foreach (string razdelitel in razdel_v)
                                    { string d = dt.Rows[strk][ii].ToString().Replace(" ", string.Empty).ToUpper();
                                        int a = d.IndexOf(razdelitel, 0);
                                        if (a > 0)
                                        {
                                            dr1[3] = d.Substring(0, a);
                                            dr1[4] = d.Substring(a, d.Length - a);
                                        }
                                    }
                                    // dr1[3] = dt.Rows[strk][ii].ToString().Replace(" ", string.Empty).ToUpper(); //prdm
                                    // если не пустые
                                    if (dr1[3].ToString()!="")
                                        dt1.Rows.Add(dr1);
                                }
                                strk++;
                           //         stroka = jj;
              //                      break;
                             //   }//break;
                            }
                        }
            dataGridView2.DataSource = dt1;
            //////
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt1.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
           // sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt1.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join("; ", fields));
            }
            textBox1.Text = sb.ToString();
            File.Delete("test.csv");
            File.WriteAllText("test.csv", sb.ToString(), Encoding.UTF8);
            //////
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
           Excel.Workbook newWorkbook = excel.Workbooks.Add();
           OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Text Files|*.xls";
            openFileDialog1.Title = "исходник";
            openFileDialog1.FileName = "от_Высших";
            //MessageBox.Show("файл с сайтами");
            string path="";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                path = openFileDialog1.FileName;
            if (path.Length > 0)            
            {
                string workbookPath = path;
                File.Delete(Path.GetTempPath()+"1234.xlsm");
                File.Delete(Path.GetTempPath() + "rozklad.csv");
                //File.Delete(Environment.SpecialFolder.Desktop + "rozklad.csv");
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(workbookPath,
                 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                excelWorkbook.SaveAs(Path.GetTempPath()+"1234", 52);
                VBIDE.VBComponent oModule;
                oModule = excelWorkbook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
                oModule.Name = "go";
                string sCode = System.IO.File.ReadAllText(@"G:\project\excell_rozklad\exel_file.vba", Encoding.Default);//.Replace("\n", " ");
                // Добавление в макрос кода .
                oModule.CodeModule.AddFromString(sCode);
                excel.Run("m");
                excelWorkbook.Close(false);
                newWorkbook.Close(false);
                excel.Quit();
                File.Delete(Path.GetTempPath()+"1234.xlsm");//лишнее удаляем
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
                File.Copy(Path.GetTempPath() + "rozklad.csv", Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
                File.Delete(Path.GetTempPath() + "rozklad.csv");//результат на раб стол
                //собираем мусор
                excelWorkbook=null; sCode = null;newWorkbook = null;
                oModule = null; excel = null;
                GC.Collect();
               
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {textBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);}

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
