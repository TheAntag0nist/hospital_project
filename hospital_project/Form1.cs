using System;
using System.Windows.Forms;
using System.IO;
using Org.BouncyCastle.Asn1.Cms;

namespace hospital_project
{
    public partial class Form2 : Form
    {
        string[] source =new string [2];
        string[] data = new string[8];
        string[] dataPrint = new string[9];
        string path = "data.txt";
        string savePath = "log.txt";
        string lastValue;

        bool errMessage = false;
        // ----------------------------------------------------------------------------
        public Form2()
        {
            InitializeComponent();
            // Set dateTimePicker to the custom format
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            // Create data.txt file if not exist
            if (!File.Exists(path))
                File.Create(path);
            // Read file and fill source 
            try
            {
                source = File.ReadAllLines(path);
                textBox1.Text = source[0];
                if (source[0] != "")
                {
                    ReadFile(source[0]);
                    remember_way.Checked = false;
                }
            }
            catch {
                if(errMessage)
                    MessageBox.Show("Error in data.txt file.");
            }
            errMessage = true;
        }
        // ----------------------------------------------------------------------------
        private void add_source_Click(object sender, EventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.DefaultExt = "*.xls;*.xlsx";
            OPF.Filter = "Microsoft Excel (*.xls*)|*.xls*|Microsoft Excel 2007 (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            OPF.Title = "Выберите документ Excel";
            OPF.Multiselect = false;
            if (OPF.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = OPF.FileName;
                source[0] = textBox1.Text;
                ReadFile(source[0]);
            }
        }
        // ----------------------------------------------------------------------------
        private void remember_way_CheckedChanged(object sender, EventArgs e)
        {
            source[0] = textBox1.Text;
            File.WriteAllLines(path, source);
        }

        private void DataAdd_Click(object sender, EventArgs e)
        {
            data[0] = dateTimePicker1.Text;
            data[1] = textBox2.Text;
            data[2] = textBox3.Text;
            data[3] = textBox4.Text;
            data[4] = textBox5.Text;
            data[5] = comboBox1.Text;
            data[6] = textBox6.Text;
            data[7] = comboBox2.Text;

            AddDataToEx(source[0],data);

            ReadFile(source[0]);
        }
        // ----------------------------------------------------------------------------
        private void Find_Click(object sender, EventArgs e)
        {
            data[0] = dateTimePicker1.Text;
            data[1] = textBox2.Text;
            data[2] = textBox3.Text;
            data[3] = textBox4.Text;
            data[4] = textBox5.Text;
            data[5] = comboBox1.Text;
            data[6] = textBox6.Text;
            data[7] = comboBox2.Text;

            int err = FindData(source[0],data);
            if (err == 1)
                MessageBox.Show("You have not filled data for searching");
        }
        // ----------------------------------------------------------------------------
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int rowNum = dataGridView1.CurrentCell.RowIndex;
            int cellNum = dataGridView1.CurrentCell.ColumnIndex;
            string currentVal = dataGridView1.Rows[rowNum].Cells[cellNum].Value.ToString();
            rowNum++;
            DialogResult saveOrNot;

            if (!File.Exists(savePath))
                File.Create(savePath);

            saveOrNot = MessageBox.Show("Хотите сохранить изменения? ", "Сохранение", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

            if (saveOrNot == DialogResult.No)
                dataGridView1.Rows[rowNum].Cells[cellNum].Value = lastValue;
            else
            {
                WriteToOneCell(source[0],rowNum,cellNum, currentVal);
                StreamWriter swLog = new StreamWriter(savePath, true);
                swLog.WriteLine(DateTime.UtcNow.ToString() + " " + source[0] + " " + " Row: " + rowNum + " Column: " + cellNum + " Value: " + lastValue);
                swLog.Close();
            }

        }
        // ----------------------------------------------------------------------------
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            lastValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString();
        }
        // ----------------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            int selectedRowCount = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);

            if (selectedRowCount > 0)
            {
                int index = dataGridView1.CurrentRow.Index;
                for (int i = 0; i < dataPrint.Length; ++i)
                    dataPrint[i] = dataGridView1.Rows[index].Cells[i].Value.ToString();

                string request = "Вы желаете распечатать справку с данными? \n\n";
                request += "ID: " + dataPrint[0] + '\n';
                request += "Дата: "+ dataPrint[1] + '\n';
                request += "ФИО плательщика: " + dataPrint[2] + '\n';
                request += "ИНН: " + dataPrint[3] + '\n';
                request += "ФИО пациента: " + dataPrint[4] + '\n';
                request += "Сумма: " + dataPrint[5] + '\n';
                request += "Год: " + dataPrint[6] + '\n';
                request += "Код услуги: " + dataPrint[7] + '\n';
                request += "Наличие чека: " + dataPrint[8] + '\n';

                DialogResult printOrNot= MessageBox.Show(request, "Печать", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

                if (printOrNot == DialogResult.Yes)
                {
                    fillTheSheat(source[1],dataPrint);
                    printMedCert(source[1]);
                }
                else
                {
                    DialogResult fillOrNot = MessageBox.Show("Заполнить справку?", "Заполнение", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if (fillOrNot == DialogResult.Yes)
                        fillTheSheat(source[1],dataPrint);
                }                 
            }
            else
                MessageBox.Show("Error: can't get selected row");
        }
    }
}
