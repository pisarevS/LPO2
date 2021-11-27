using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Учёт_колёс
{
    internal class _Excel
    {
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Workbook workbookTemplate;
        private Excel.Worksheet worksheet;
        private _Close _close = new _Close();
        private StreamReader sr;
        private StreamWriter sw;
        private int iLastRow;

        public void UploadInExcel(Excel.Application ExcelApp, DataGridView dataGridView1, SaveFileDialog saveFileDialog1)
        {
            int namber, namberDouble, column = 0;
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    namber = Convert.ToInt32(dataGridView1[1, i].Value.ToString());
                    namberDouble = namber;
                    if (namber <= 55)
                    {
                        column = 2;
                    }
                    if (namber >= 56 && namber <= 111)
                    {
                        namber -= 56;
                        column = 4;
                    }
                    if (namber >= 112 && namber <= 167)
                    {
                        namber -= 112;
                        column = 6;
                    }
                    if (namber >= 168 && namber <= 223)
                    {
                        namber -= 168;
                        column = 8;
                    }
                    if (namber >= 224 && namber <= 279)
                    {
                        namber -= 224;
                        column = 10;
                    }
                    if (namber >= 280 && namber <= 335)
                    {
                        namber -= 280;
                        column = 12;
                    }
                    if (namber >= 336 && namber <= 391)
                    {
                        namber -= 336;
                        column = 14;
                    }
                    if (namber >= 392 && namber <= 447)
                    {
                        namber -= 392;
                        column = 16;
                    }
                    if (namber >= 448 && namber <= 499)
                    {
                        namber -= 448;
                        column = 18;
                    }
                    if (namber >= 500 && namber <= 555)
                    {
                        namber -= 500;
                        column = 20;
                    }
                    if (namber >= 556 && namber <= 611)
                    {
                        namber -= 556;
                        column = 22;
                    }
                    if (namber >= 612 && namber <= 667)
                    {
                        namber -= 612;
                        column = 24;
                    }
                    if (namber >= 668 && namber <= 723)
                    {
                        namber -= 668;
                        column = 26;
                    }
                    if (namber >= 724 && namber <= 779)
                    {
                        namber -= 724;
                        column = 28;
                    }
                    if (namber >= 780 && namber <= 835)
                    {
                        namber -= 780;
                        column = 30;
                    }
                    if (namber >= 836 && namber <= 891)
                    {
                        namber -= 836;
                        column = 32;
                    }
                    if (namber >= 892 && namber <= 947)
                    {
                        namber -= 892;
                        column = 34;
                    }
                    if (namber >= 948 && namber <= 999)
                    {
                        namber -= 948;
                        column = 36;
                    }

                    if (ExcelApp.Cells[namber + 4, column].Value2 == null)
                    {
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.A) ExcelApp.Cells[namber + 4, column].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.B) ExcelApp.Cells[namber + 4, column].Font.Color = System.Drawing.Color.FromArgb(153, 51, 255);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.V) ExcelApp.Cells[namber + 4, column].Font.Color = System.Drawing.Color.FromArgb(0, 204, 64);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.G) ExcelApp.Cells[namber + 4, column].Font.Color = Excel.XlRgbColor.rgbBlack;

                        ExcelApp.Cells[namber + 4, column] = dataGridView1.Rows[i].Cells[2].Value.ToString();                  
                     }
                    else
                    {
                        if(ExcelApp.Cells[2, 39].Value == null) {
                            Excel.Range _excelCells2 = ExcelApp.get_Range("AM2", "AN2").Cells;
                            ExcelApp.Cells[2, 39] = "Двойники";
                            _excelCells2.Merge(Type.Missing);
                            ExcelApp.Cells[3, 39] = "№";
                            ExcelApp.Cells[3, 40] = "Станок";
                            _excelCells2.HorizontalAlignment =Excel.Constants.xlCenter;                         
                        }

                        iLastRow = worksheet.Cells[worksheet.Rows.Count, 39].End[Excel.XlDirection.xlUp].Row;                      
                        ExcelApp.Cells[iLastRow + 1, 39] = namberDouble;
                        ExcelApp.Cells[iLastRow + 1, 40] = dataGridView1.Rows[i].Cells[2].Value.ToString();

                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.A) ExcelApp.Cells[iLastRow + 1, 40].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.B) ExcelApp.Cells[iLastRow + 1, 40].Font.Color = System.Drawing.Color.FromArgb(153, 51, 255);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.V) ExcelApp.Cells[iLastRow + 1, 40].Font.Color = System.Drawing.Color.FromArgb(0, 204, 64);
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == Const.G) ExcelApp.Cells[iLastRow + 1, 40].Font.Color = Excel.XlRgbColor.rgbBlack;
                    }
                }
            }
            catch
            {
                try
                {               
                    ExcelApp.ActiveWorkbook.ActiveSheet(saveFileDialog1.FileName.ToString());
                    ExcelApp.ActiveWorkbook.ActiveSheet.Saved = true;
                    ExcelApp.Quit();
                    GC.Collect();
                }
                catch
                {
                    ExcelApp.Quit();
                    GC.Collect();
                }
            }
        }

        public void CreateExcel(string path, SaveFileDialog saveFileDialog1)
        {
            excelApp = new Excel.Application();          
            saveFileDialog1.Title = "Создать";
            saveFileDialog1.InitialDirectory = @"D:\";
            saveFileDialog1.Filter = "Книга Excel (*.xlsx)|*.xlsx|Книга Excel 97-2003 (*.xls)|*.xls|Все файлы (*.*)|*.*";
            saveFileDialog1.ShowDialog();
            string filename = saveFileDialog1.FileName;
            try
            {
                if (filename != "")
                {
                    sw = new StreamWriter(path, false, Encoding.Default);
                    excelApp.SheetsInNewWorkbook = 1;
                    workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filename);
                    filename = saveFileDialog1.FileName;
                    sw.Write(filename);
                    sw.Close();
                    excelApp.Quit();
                }
            }
            catch { }
            filename = "";
            saveFileDialog1.FileName = "";
        }

        public void OpenExcel(string path, OpenFileDialog openFileDialog1)
        {
            openFileDialog1.ShowDialog();
            openFileDialog1.InitialDirectory = @"D:\";
            string filename = openFileDialog1.FileName;
            if (filename == "")
            {
                sr = new StreamReader(path, Encoding.Default);
                filename = sr.ReadLine();
                sr.Close();
            }
            else
            {
                filename = openFileDialog1.FileName;
                for (int i = 0; i < filename.Length; i++)
                {
                    if (filename[i] == '.')
                    {
                        for (int j = i; j < filename.Length; j++)
                        {
                            filename = filename.Remove(i);
                        }

                    }
                }
                sw = new StreamWriter(path, false, Encoding.Default);
                sw.Write(filename);
                sw.Close();
            }
            filename = "";
            openFileDialog1.FileName = "";
        }

        public void ExportToExcel(string path, string template, TextBox textBox1, DataGridView dataGridView1, SaveFileDialog saveFileDialog1)
        {
            _close.CloseProcess();
            string worksheetname = textBox1.Text;
            sr = new StreamReader(path, Encoding.Default);
            string filename = sr.ReadLine();
            sr.Close();
            excelApp = new Excel.Application();
           
            try
            {
                workbook = excelApp.Workbooks.Open(filename);
            }
            catch
            {
                excelApp.Quit();
                GC.Collect();
                _close.CloseProcess();
                MessageBox.Show("Файл не найден :(");
                return;
            }

            excelApp.Visible = false;
            worksheet = null;
            int sheetscount = workbook.Sheets.Count;
            if (ShExist(workbook, worksheetname) == false)
            {
                //Открываем файл шаблона Excel
                try
                {
                    workbookTemplate = excelApp.Workbooks.Open(template);
                }
                catch
                {
                    excelApp.Quit();
                    GC.Collect();
                    _close.CloseProcess();
                    MessageBox.Show("Файл шаблона не найден :(");
                    return;
                }
                //Создаем лист Excel, если лист с таким именем отсутствует
                try
                {
                    excelApp.Visible = false;
                    worksheet = (Excel.Worksheet)workbookTemplate.Sheets[1];
                    worksheet.Name = worksheetname;
                    workbookTemplate.Save();
                    workbookTemplate.Worksheets[1].Copy(After: workbook.Worksheets[sheetscount]);
                    excelApp.Visible = true;
                }
                catch {}
                //Закрываем шаблон файла Excel
                excelApp.WindowState = Excel.XlWindowState.xlMaximized;
                workbookTemplate.Close(true);              
                worksheet = (Excel.Worksheet)workbook.Sheets.get_Item(worksheetname);
                //Выгрузка массива в лист Excel
                UploadInExcel(excelApp, dataGridView1, saveFileDialog1);
                return;
            }
            else
            {
                excelApp.WindowState = Excel.XlWindowState.xlMaximized;
                excelApp.Visible = true;
                worksheet = (Excel.Worksheet)workbook.Sheets.get_Item(worksheetname);
                worksheet.Activate();
                //Выгрузка массива в лист Excel
                UploadInExcel(excelApp, dataGridView1, saveFileDialog1);
            }
        }

        public static bool ShExist(object workbook, string worksheetname)
        {
            bool bEx = false;
            object wsPShts = null, wsSh = null;
            wsPShts = workbook.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, workbook, null);
            try
            {
                wsSh = wsPShts.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, wsPShts, new object[] { worksheetname });
                bEx = true;
            }
            catch { bEx = false; }
            return (bEx);
        }
    }
}
