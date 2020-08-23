using System.Data;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data.OleDb;
/*
Project use  Npoi (for .xls/.xlsx) 
and ExcelDataReader (for reading .xls,.xlsx).
In the future will be used only Interop.Excel;
*/


namespace hospital_project
{
    public partial class Form2 : Form
    {
        static HSSFWorkbook hssfworkbook;
        static XSSFWorkbook xssfworkbook;
        static int lastRow;
        static int lastCol;

        static bool startOrNot = true;
        static bool addData = false;

        // Show data from file
        void ReadFile(string filePath)
        {
            bool xlsx;

            if (filePath.Contains("xlsx"))
                xlsx = true;
            else
                xlsx = false;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // The result of each spreadsheet is in result.Tables
                    var result = reader.AsDataSet();
                    int i = 0;
                    // Set count of column and rows
                    while (true)
                    {
                        var dataTable = result.Tables[i];

                        if (dataTable.TableName == "Реестр")
                        {
                            if (dataTable.Columns.Count > 0 && dataTable.Rows.Count > 0)
                            {
                                SetRowsCols(dataTable, dataGridView1, xlsx);
                                dataGridFill(dataTable);
                            }
                            break;
                        }
                        else
                        {
                            ++i;
                            try
                            {
                                dataTable = result.Tables[i];
                            }
                            catch
                            {
                                MessageBox.Show("Not found Реестр and ends of sheets");
                                dataTable = result.Tables[0];
                                if (dataTable.Columns.Count > 0 && dataTable.Rows.Count > 0)
                                {
                                    SetRowsCols(dataTable, dataGridView1, xlsx);
                                    dataGridFill(dataTable);
                                }
                                break;
                            }
                        }
                    }
                }
            }

            if (startOrNot)
            {
                dataGridView1.Rows[0].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                startOrNot = false;
            }
            else if(addData)
            {
                dataGridView1.Rows[lastRow-2].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[lastRow-2].Cells[0];
                addData = false;
            }

        }

        // Sets rows and columns for dataGridView
        void SetRowsCols(DataTable dataTable, DataGridView dataGridView1, bool xlsxFile)
        {
            dataGridView1.ColumnCount = dataTable.Columns.Count;
            dataGridView1.RowCount = dataTable.Rows.Count;
            lastRow = dataTable.Rows.Count;
            lastCol = dataTable.Columns.Count;
        }
        // Fills the dataGridView
        void dataGridFill(DataTable dataTable)
        {
            // Fill the headers columns
            for (int j = 0; j < dataTable.Columns.Count; ++j)
            {
                dataGridView1.Columns[j].HeaderCell.Value = dataTable.Rows[0][j];
            }

            // Show data in dataGridView
            for (int i = 1; i < dataTable.Rows.Count; ++i)
            {
                for (int j = 0; j < dataTable.Columns.Count; ++j)
                {
                    dataGridView1.Rows[i - 1].Cells[j].Value = dataTable.Rows[i][j];
                }
            }
        }

        // Works for .xls/.xlsx files
        static void AddDataToEx(string filePath, string[] data)
        {
            addData = true;
            InitializeWorkbook(filePath);

            ICell cell;
            ISheet sheet = null;

            if (filePath.Contains("xlsx"))
            {
                try
                {
                    int i = 0;
                    while (true)
                    {
                        sheet = xssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
                catch
                {
                    sheet = xssfworkbook.GetSheetAt(0);
                }
            }
            else if (filePath.Contains("xls"))
            {
                try
                {
                    int i = 0;
                    while (true)
                    {
                        sheet = hssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
                catch
                {
                    sheet = hssfworkbook.GetSheetAt(0);
                }
            }


            IRow row = sheet.CreateRow(lastRow);

            cell = row.CreateCell(0);

            cell.SetCellValue(lastRow);

            for (int i = 1; i < lastCol; ++i)
            {
                cell = row.CreateCell(i);
                cell.SetCellValue(data[i - 1]);
            }

            ++lastCol;

            WriteToFile(filePath);
        }

        static void WriteToFile(string filePath)
        {
            //Write the stream data of workbook 
            FileStream file = new FileStream(filePath, FileMode.Open);
            if (filePath.Contains("xlsx"))
                xssfworkbook.Write(file);
            else if (filePath.Contains("xls"))
                hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook(string filePath)
        {
            try
            {
                using (var fs = File.OpenRead(filePath))
                {
                    if (filePath.Contains("xlsx"))
                        xssfworkbook = new XSSFWorkbook(fs);
                    else if (filePath.Contains("xls"))
                        hssfworkbook = new HSSFWorkbook(fs);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel read error");
                return;
            }
        }

        int FindData(string filePath, string[] data)
        {
            InitializeWorkbook(filePath);

            // Counter of null elements in data array
            int cntNull = 0;

            for (int i = 1; i < data.Length; ++i)
            {
                if (data[i] == null || data[i] == "")
                    ++cntNull;
            }
            // Return error
            if (cntNull == data.Length - 1)
                return 1;

            ISheet sheet = null;
            try
            {
                int i = 0;
                if (filePath.Contains("xlsx"))
                {
                    while (true)
                    {
                        sheet = xssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
                else if (filePath.Contains("xls"))
                {
                    while (true)
                    {
                        sheet = hssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
            }
            catch
            {
                if (filePath.Contains("xlsx"))
                    sheet = xssfworkbook.GetSheetAt(0);
                else if (filePath.Contains("xls"))
                    sheet = hssfworkbook.GetSheetAt(0);
            }

            if (simpleSearch(sheet, data))
                MessageBox.Show("Not found");

            return 0;
        }

        bool simpleSearch(ISheet sheet, string[] data)
        {
            byte check = 0,
                 temp = 0,
                 plus = 2;

            ICell cell;

            for (int i = 1; i < data.Length; ++i)
            {
                if (data[i] != "" && data[i] != null)
                    check += plus;
                plus <<= 1;

            }

            for (int j = 1; j < lastRow; ++j)
            {
                plus = 2;

                for (int i = 1; i < lastCol - 1; ++i)
                {
                    try
                    {
                        cell = sheet.GetRow(j).GetCell(i + 1);
                    }
                    catch
                    {
                        continue;
                    }

                    if (cell.CellType == CellType.String)
                    {
                        if (cell.StringCellValue == data[i] && data[i] != "" && data[i] != null)
                            temp += plus;
                        plus <<= 1;
                    }
                    else if (cell.CellType == CellType.Numeric && data[i] != "" && data[i] != null)
                    {
                        if (cell.NumericCellValue.ToString() == data[i])
                            temp += plus;
                        plus <<= 1;
                    }
                    else
                        plus <<= 1;
                }

                DialogResult endSearch = DialogResult.Yes;

                if (temp == check)
                {
                    dataGridView1.Rows[j - 1].Selected = true;
                    dataGridView1.CurrentCell = dataGridView1.Rows[j - 1].Cells[0];
                    endSearch = MessageBox.Show("We found that on line " + j + "\nFind next?\n", "Continue", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    dataGridView1.ClearSelection();
                }

                if (endSearch == DialogResult.No)
                {
                    dataGridView1.Rows[j-1].Selected = true;
                    dataGridView1.CurrentCell = dataGridView1.Rows[j - 1].Cells[0];
                    return false;
                }

                temp = 0x00;
            }
            return true;
        }

        void WriteToOneCell(string filePath, int rowNum, int cellNum, string data, int sheetNum = 0)
        {
            InitializeWorkbook(filePath);

            ICell cell = null;
            ISheet sheet = null;

            if (filePath.Contains("xlsx"))
            {
                try
                {
                    int i = 0;
                    while (true)
                    {
                        sheet = xssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
                catch
                {
                    sheet = xssfworkbook.GetSheetAt(sheetNum);
                }
            }
            else if (filePath.Contains("xls"))
            {
                try
                {
                    int i = 0;
                    while (true)
                    {
                        sheet = hssfworkbook.GetSheetAt(i);
                        if (sheet.SheetName == "Реестр")
                            break;
                        ++i;
                    }
                }
                catch
                {
                    sheet = hssfworkbook.GetSheetAt(sheetNum);
                }
            }
            try
            {
                cell = sheet.GetRow(rowNum).GetCell(cellNum);
            }
            catch
            {
                if (cell == null)
                {
                    IRow row = sheet.CreateRow(rowNum);
                    cell = row.CreateCell(cellNum);
                }
            }

            cell.SetCellValue(data);

            WriteToFile(filePath);
        }

        void fillTheSheat(string filePath, string[] data)
        {
            for (int i = 0; i < data.Length; ++i)
                WriteToOneCell(filePath,i+1,1,data[i]);
        }

        void printMedCert(string filePath)
        {
            var fileName = filePath;
            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                    var adapter = new OleDbDataAdapter(cmd);
                    var ds = new DataSet();
                    adapter.Fill(ds);
                }
            }


        }
    }
}
