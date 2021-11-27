using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Учёт_колёс
{
    public partial class Form1 : Form
    {
        private DataGridViewTextBoxCell txtxCell;
        private StreamReader sr;
        private readonly _Excel _excel = new _Excel();
        private readonly string path = Directory.GetCurrentDirectory();
        private readonly string template = "\\config\\template.xls";
        private readonly string lasteFileOpen = "\\config\\path.txt";
        private readonly string machine_list = "\\config\\machine_list.txt";
        private readonly string tab = "\\config\\tab.txt";
        private int namber = 1;

        public Form1()
        {
            template = path + template;
            lasteFileOpen = path + lasteFileOpen;
            machine_list = path + machine_list;
            tab = path + tab;
            InitializeComponent();
            radioButton1.Checked = true;

            try
            {
                sr = new StreamReader(tab, Encoding.Default);
                tabNumber.Text = sr.ReadLine();
                sr.Close();

                sr = new StreamReader(lasteFileOpen, Encoding.Default);
                Text = sr.ReadLine();
                sr.Close();
            }
            catch { }
            openFileDialog1.FileName = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                sr = new StreamReader(machine_list, Encoding.Default);
                while (!sr.EndOfStream)
                    machineNumber.Items.Add(sr.ReadLine());
                sr.Close();
            }
            catch { }
            Application.UseWaitCursor = false;
        }

        public void EnabledButton()
        {
            if (machineNumber.Text != "" && tabNumber.Text != "" && wheelNumber.Text != "")
            {
                enter.Enabled = true;
            }
            else
            {
                enter.Enabled = false;
            }
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _excel.OpenExcel(lasteFileOpen, openFileDialog1);
            sr = new StreamReader(lasteFileOpen, Encoding.Default);
            Text = sr.ReadLine();
            sr.Close();
        }

        private void CreateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _excel.CreateExcel(lasteFileOpen, saveFileDialog1);
            sr = new StreamReader(lasteFileOpen, Encoding.Default);
            Text = sr.ReadLine();
            sr.Close();
        }

        private void EnterButton(object sender, EventArgs e)
        {
            string workShift = "";
            if (radioButton1.Checked) workShift = Const.A;
            if (radioButton2.Checked) workShift = Const.B; 
            if (radioButton3.Checked) workShift = Const.V; 
            if (radioButton4.Checked) workShift = Const.G;

            dataGridView1.Rows.Add(Convert.ToString(namber++), wheelNumber.Text, machineNumber.Text,workShift);
            wheelNumber.Text = "";
            if (dataGridView1.Rows.Count != 0) save.Enabled = true;

            dataGridView1.ClearSelection();
            int index=dataGridView1.RowCount;
            if (index >= 2)
            {
                dataGridView1.Rows[dataGridView1.RowCount - 2].Selected = true;
            }
            else
            {
                dataGridView1.Rows[dataGridView1.RowCount - 1].Selected = true;
            }
            


            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;

        }

        private void DeleteAllButton(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            namber = 1;
            if (dataGridView1.Rows.Count == 1) save.Enabled = false;
        }

        private void DeleteCellButton(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount >= 2)
            {
                int indexSellect = dataGridView1.SelectedCells[0].RowIndex;
                int count = dataGridView1.RowCount;
                count -= 2;
                dataGridView1.Rows.RemoveAt(indexSellect);
                for (int i = 0; i < count; i++)
                {
                    txtxCell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[0];
                    txtxCell.Value = Convert.ToString(i + 1);
                }
                namber = count + 1;
                if (dataGridView1.Rows.Count == 1)
                {
                    save.Enabled = false;
                }
            }
        }

        private void EnterButtonKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && wheelNumber.Text != ""&& machineNumber.Text != "" && tabNumber.Text != "")
            {
                EnterButton(sender, e);
                e.SuppressKeyPress = true;
            }
            if (e.KeyCode == Keys.F5 && save.Enabled == true) SaveButton(sender, e);
        }

        private void SaveButton(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount >= 2) _excel.ExportToExcel(lasteFileOpen, template, tabNumber, dataGridView1, saveFileDialog1);
        }

        private void SaveToolStripMenuItemClick(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount >= 2) _excel.ExportToExcel(lasteFileOpen, template, tabNumber, dataGridView1, saveFileDialog1);
        }

        private void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
        {
            EnabledButton();
        }

        private void MeltingNumberTextChanged(object sender, EventArgs e)
        {
            EnabledButton();
        }

        private void WheelNumberTextChanged(object sender, EventArgs e)
        {
            EnabledButton();
        }

        private void ExitToolStripMenuItemClick(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            StreamWriter sw;
            sw = new StreamWriter(tab, false, Encoding.Default);
            sw.Write(tabNumber.Text);
            sw.Close();
        }
    }
}
