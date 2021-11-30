using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace dstv_qtt_change
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Filter = "xls files(*.xls|*.xls|*.xlsx|*.xlsx|All files(*.*)|*.*";
        }

        private void button1_Click(object sender, EventArgs e) //openExel
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            
            string filename = openFileDialog1.FileName;// получаем выбранный файл            
            textBox1.Text = filename; //путь файла в тестовое поле           
        }

        private void pathToDSTV_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = folderBrowserDialog1.SelectedPath;
            textBox2.Text = filename; //путь файла в тестовое поле   
        }

        private void Start_Click(object sender, EventArgs e)
        {
            try
            {
                Dictionary<string, int> Pos = GetDictionaryOfPartsNonOffice(textBox1.Text);
                //Dictionary<int, int> Pos1 = GetDictionaryOfPartsOfice(textBox1.Text);
                  //GetAllDSTV(textBox2.Text);
                 changeDSTV(Pos, GetAllDSTV(textBox2.Text));
                 //MessageBox.Show("Выполнение завершено");
            }
            catch (Exception ex)
            {
                messages.Text ="Error: " + ex.Source +  ex.Message;
            }
                    
        }
        private Dictionary<string, int> GetDictionaryOfPartsNonOffice(string pathToExcel)
        {
            FileStream xlsFile = new FileStream(pathToExcel, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hssfwb = new HSSFWorkbook(xlsFile);
            //ISheet sheet1 = hssfwb.GetSheet("_5_Список деталей");
            HSSFSheet sheet = (HSSFSheet)hssfwb.GetSheetAt(0);
            Dictionary<string, int> PositionQtt = new Dictionary<string, int>();
            for (int row = 3; row <= (sheet.LastRowNum-1); row++)
            {
                //получаем строку
                string position = sheet.GetRow(row).Cells[1].ToString();
                int qtt = Convert.ToInt32(sheet.GetRow(row).Cells[3].ToString());
                PositionQtt.Add(position, qtt);
            }
            return PositionQtt;

        }
        private Dictionary<int, int> GetDictionaryOfPartsOfice(string pathToExcel)
        {
           // int startLine = 4; // начинаем считывает со строки 4(в ексель А=4,В=4... и т.д.)
           // char column1 = 'B'; //считываемый столбец
           // char column2 = 'D'; //считываемый столбец

            Dictionary<int, int> PositionQtt = new Dictionary<int, int>();

            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@pathToExcel, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            //var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            //int iLastRow = ObjWorkSheet.Cells[ObjWorkSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            //var arrData = (object[,])ObjWorkSheet.Range["B:D" + iLastRow].Value;
            //var x = ObjWorkSheet.Range["B"];

            Excel.Range columnB = ObjWorkSheet.UsedRange.Columns[2];
            Excel.Range columnD = ObjWorkSheet.UsedRange.Columns[4];
            System.Array valB = (System.Array)columnB.Cells.Value;
            System.Array valD = (System.Array)columnD.Cells.Value;
            string[] position = valB.OfType<object>().Select(o => o.ToString()).ToArray();
            string[] qtt = valD.OfType<object>().Select(o => o.ToString()).ToArray();
            //Удаляем текстовое поле
            position = position.Where(val => val != "Позиція").ToArray();
            qtt = qtt.Where(val => val != "Кількість, шт").ToArray();
            for (int i = 0; i < position.Length; i++)
            {
                PositionQtt.Add(Convert.ToInt32(position[i]), Convert.ToInt32(qtt[i]));
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя

            return PositionQtt;
        }

        private string[] GetAllDSTV(string path)
        {
            string[] files = Directory.GetFiles($@"{path}", "*.nc1", SearchOption.AllDirectories);
            //this.listBox1.Items.AddRange(files);
            //int x = 34;
            return files;
        }

        private void changeDSTV(Dictionary<string, int> excelDatas, string[] pathsToDSTV)
        {
            List<string> messagesErr = new List<string>();
            messagesErr.Add("Выполнение завершено");
           foreach (KeyValuePair <string, int> key in excelDatas)
            {
                if (searchFile(key.Key, pathsToDSTV))
                {
                    IEnumerable<string> pathToDSTV = pathsToDSTV.Where(val => val == (Path.GetDirectoryName(val) +"\\" + $"{key.Key}.nc1"));
                   if(pathToDSTV.ToArray().Length == 1)
                    {
                        string[] path = pathToDSTV.ToArray();
                        changeQttNC(path[0], key.Value);
                    }
                   else
                    {
                        string[] paths = pathToDSTV.ToArray();
                        messagesErr.Add("Одинаковый номер дств в разных папках...");
                        foreach (string val in paths)
                        {
                            messagesErr.Add(val);
                        }
                        
                    }
                }
                else
                {
                    messagesErr.Add("Файл дств не найден:" + key.Key);

                }
            }
            string[] messagesErrRes = messagesErr.ToArray();
            string messER = String.Join("\r\n", messagesErrRes);

            messages.Text = messER;
        }

        private bool searchFile(string position, string[] pathsToDSTV)
        {            
             foreach (string path in pathsToDSTV)
             {
                string dstv = Path.GetDirectoryName(path) + "\\" + $"{position}.nc1";
                // string[] words = path.Split(new char[] { '\\' });
                // string dstvFileName = words[words.Length-1];
                if (dstv == path)
                 {
                     return true;
                 }                 
             }
            return false;            
            /*foreach (string path in pathsToDSTV)
            {
                var x = Path.GetDirectoryName(path)+"\\" + $"{position}.nc1";
            }*/
            //IEnumerable<string> result = pathsToDSTV.Where(val => val == (Path.GetDirectoryName(val) +"\\" + $"{position}.nc1"));
           
        }    
        
        private void changeQttNC (string pathToDSTV, int qtt)
        {
            try
            {
                StreamReader sr = new StreamReader(pathToDSTV, System.Text.Encoding.Default);
                var datasSR = sr.ReadToEnd();
                string[] words = datasSR.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int qttStrFromDstv = Convert.ToInt32((words[7].Trim(new char[] { ' ', '\r' })));    
                if (qtt != qttStrFromDstv) 
                {
                    words[7] = "  " + qtt + "\r";
                }
                else
                {
                    return;
                }

                for (int i = 0; i<words.Length; i++)
                {
                    words[i] = words[i] + '\n';
                }
                sr.Close();
                ///перезапись файла дств
                string newDstvData = String.Join("", words);
                StreamWriter sw = new StreamWriter(pathToDSTV, false, System.Text.Encoding.Default);
                sw.WriteLine(newDstvData);
                sw.Close();



            }
            catch (Exception e)
            {

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void messages_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
