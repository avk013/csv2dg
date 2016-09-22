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
using System.Net;
using System.Collections.Specialized;

namespace csv2dg
{
    public partial class Form1 : Form
    {
        //        private int applicationHwnd;
        //      private object ExcelApplication;
        string file_path= @"G:\project\excell_rozklad\";
        public Form1()
        {
            InitializeComponent();
        }
        public static string UploadFileEx(string uploadfile, string url,
           string fileFormName, string contenttype, NameValueCollection querystring,
           CookieContainer cookies, TextBox ino=null, TextBox outo=null)//процедура загрузки
        {if ((fileFormName == null) || (fileFormName.Length == 0))
            {fileFormName = "file";}
         if ((contenttype == null) || (contenttype.Length == 0))
            {contenttype = "application/octet-stream";}
            string postdata;
            postdata = "?";
            if (querystring != null)
            {foreach (string key in querystring.Keys)
                {postdata += key + "=" + querystring.Get(key) + "&";}}
            Uri uri = new Uri(url + postdata);
            string boundary = "----------" + DateTime.Now.Ticks.ToString("x");
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(uri);
            webrequest.CookieContainer = cookies;
            webrequest.ContentType = "multipart/form-data; boundary=" + boundary;
            webrequest.Method = "POST";
            // Build up the post message header
            StringBuilder sb = new StringBuilder();
            sb.Append("--");sb.Append(boundary);sb.Append("\r\n");
            sb.Append("Content-Disposition: form-data; name=\"");
            sb.Append(fileFormName);sb.Append("\"; filename=\"");
            //sb.Append(Path.GetFileName(uploadfile));
            sb.Append(Path.GetFileName(uploadfile).Substring(0, (Path.GetFileName(uploadfile).Length-4)));
            sb.Append("\"");sb.Append("\r\n");sb.Append("Content-Type: ");
            sb.Append(contenttype);sb.Append("\r\n");sb.Append("\r\n");
            string postHeader = sb.ToString();
            if (ino!=null) ino.Text = url + postdata+postHeader;
            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(postHeader);
            // Build the trailing boundary string as a byte array
            // ensuring the boundary appears on a line by itself
            byte[] boundaryBytes =
                   Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");
            FileStream fileStream = new FileStream(uploadfile,
                                        FileMode.Open, FileAccess.Read);
            long length = postHeaderBytes.Length + fileStream.Length +
                                                   boundaryBytes.Length;
            webrequest.ContentLength = length;
            Stream requestStream = webrequest.GetRequestStream();
            // Write out our post header
            requestStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
            // Write out the file contents
            byte[] buffer = new Byte[checked((uint)Math.Min(4096,
                                     (int)fileStream.Length))];
            int bytesRead = 0;
            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                requestStream.Write(buffer, 0, bytesRead);
            // Write out the trailing boundary
            requestStream.Write(boundaryBytes, 0, boundaryBytes.Length);
            WebResponse responce = webrequest.GetResponse();
            Stream s = responce.GetResponseStream();
            StreamReader sr = new StreamReader(s);
            return sr.ReadToEnd(); }

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
                                } strk++;}}
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
            File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
            //File.Copy(Path.GetTempPath() + "rozklad.csv", Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
            //File.Delete("test.csv");
            //File.WriteAllText("test.csv", sb.ToString(), Encoding.UTF8);
            File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv", sb.ToString(), Encoding.UTF8);
            //////
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
           //Excel.Workbook newWorkbook = excel.Workbooks.Add();
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
               ////File.Delete(Environment.SpecialFolder.Desktop + "rozklad.csv");
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(workbookPath,
                 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                excelWorkbook.SaveAs(Path.GetTempPath()+"1234", 52);
                VBIDE.VBComponent oModule;
                oModule = excelWorkbook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
                oModule.Name = "go";
    //string sCode = System.IO.File.ReadAllText(file_path+"exel_file.vba", Encoding.Default);//.Replace("\n", " ");
   //or
                string sCode = "Sub ununion() ' объединенные области разъединяем и заполняем \n"+
    "Dim rCell As Range, sValue$, sAddress$, i& \n" +
    "Application.ScreenUpdating = False\n" +
      "Set iSource = [A1: L600]\n" +
       " For Each rCell In iSource\n" +
    "If rCell.MergeCells Then\n" +
     "   sAddress = rCell.MergeArea.Address: rCell.UnMerge\n" +
      "  Range(sAddress).Value = rCell.Value\n" +
    "End If\n" +
    "Next\n" +
    "Application.ScreenUpdating = True\n" +
"End Sub\n" +
"Sub m()\n" +
"Dim name As String\n" +
"name = \"result\"\n" +
"Dim oSheet As Excel.Worksheet\n" +
"Set oSheet = Worksheets.Add()\n" +
"oSheet.name = name\n" +
 "   For i = 1 To Sheets.Count 'перебираем все листы\n" +
  "      If Sheets(i).name <> name Then\n" +
   "        myR_Total = Sheets(name).Range(\"A\" & Sheets(name).Rows.Count).End(xlUp).Row\n" +
    "       myR_i = Sheets(i).Range(\"A\" & Sheets(i).Rows.Count).End(xlUp).Row\n" +
     "      Sheets(i).Rows(\"1:\" & myR_i).Copy Destination:= Sheets(name).Range(\"A\" & myR_Total + 1)\n" +
      "  End If\n" +
    "Next\n" +
 "Set ws = ActiveWorkbook.Sheets(name)\n" +
 "ws.Activate\n" +
  "  ununion\n" +
"Dim r As Long, rng As Range ' удаляем пустые строки\n" +
    "For r = 1 To ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count\n" +
        "If Application.CountA(Rows(r)) = 0 Then\n" +
         "   If rng Is Nothing Then Set rng = Rows(r) Else Set rng = Union(rng, Rows(r))\n" +
        "End If\n" +
    "Next r\n" +
    "If Not rng Is Nothing Then rng.Delete\n" +
"ActiveSheet.Copy\n" +
"'Kill(\"rozklad.csv\") \n"+
"ActiveWorkbook.SaveAs ThisWorkbook.Path & \"\\\" & \"rozklad.csv\", xlCSV, CreateBackup:=False, Local:=True\n" +
"ActiveWorkbook.Close 0\n" +
"End Sub";
                // Добавление в макрос кода .
                oModule.CodeModule.AddFromString(sCode);
                excel.Run("m");
                excelWorkbook.Close(false);
             ////   newWorkbook.Close(false);
                excel.Quit();
                File.Delete(Path.GetTempPath()+"1234.xlsm");//лишнее удаляем
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
                File.Copy(Path.GetTempPath() + "rozklad.csv", Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv");
                File.Delete(Path.GetTempPath() + "rozklad.csv");//результат на раб стол
                //собираем мусор
                excelWorkbook=null; sCode = null;
                ////newWorkbook = null;
                oModule = null; excel = null;
                GC.Collect();}}
        private void button3_Click(object sender, EventArgs e)
        {textBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);}

        private void Form1_Load(object sender, EventArgs e)
        {}
      
        private void button4_Click(object sender, EventArgs e)
        {//вызов загрузки файла на сервер
  CookieContainer cookies = new CookieContainer();
        //add or use cookies
        NameValueCollection querystring = new NameValueCollection();
        querystring["uname"]="uname";
querystring["passwd"]="rozklad";
string uploadfile;// set to file to upload
        uploadfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\rozklad.csv";
//everything except upload file and url can be left blank if needed
string outdata = UploadFileEx(uploadfile,
     "http://fei.idgu.edu.ua/rozklad+/server/file.php", "uploadfile", "image/pjpeg",
     querystring, cookies, textBox2);
     textBox1.Text = outdata;
        } 
    }
}
