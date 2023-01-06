using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace ReaderEngine
{
    public class Helper
    {
        ConnectionDB connectionDB = new ConnectionDB();

        // for read excel file
        public System.Data.DataTable ReadExcel(string fileName, string fileExt, string query)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HRD=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter(query, con); //here we read data from sheet1                                                       
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }

        public String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }


                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }


        //for encrypt password
        public string encryption(String password)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] encrypt;
            UTF8Encoding encode = new UTF8Encoding();
            //encrypt the given password string into Encrypted data  
            encrypt = md5.ComputeHash(encode.GetBytes(password));
            StringBuilder encryptdata = new StringBuilder();
            //Create a new string by using the encrypted data  
            for (int i = 0; i < encrypt.Length; i++)
            {
                encryptdata.Append(encrypt[i].ToString());
            }
            return encryptdata.ToString();
        }

        //get root treeview
        public TreeNode FindRootNode(TreeNode treeNode)
        {
            while (treeNode.Parent != null)
            {
                treeNode = treeNode.Parent;
            }
            return treeNode;
        }

        public bool IsTheSameCellValue(DataGridView dataGridView, int column, int row)
        {
            DataGridViewCell cell1 = dataGridView[column, row];
            DataGridViewCell cell2 = dataGridView[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }
            return cell1.Value.ToString() == cell2.Value.ToString();
        }

        public string randomText(int length)
        {
            var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var Charsarr = new char[length];
            var random = new Random();

            for (int i = 0; i < Charsarr.Length; i++)
            {
                Charsarr[i] = characters[random.Next(characters.Length)];
            }

            var resultString = new String(Charsarr);
            return resultString;
        }

        // to fill listbox from db
        public void fill_listbox(string sql, ListBox lst, string column)
        {
            try
            {
                lst.DataSource = null;
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(sql, connectionDB.connection))
                {
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        lst.DataSource = dt;
                        lst.DisplayMember = column;
                        lst.ValueMember = column;
                    }
                    else
                    {
                        lst.Items.Add("No Data");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // to fill checklistbox from db
        public void fill_checklistbox(string sql, CheckedListBox lst, string column)
        {
            try
            {
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(sql, connectionDB.connection))
                {
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        lst.DataSource = dt;
                        lst.DisplayMember = column;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // to fill dgv from db
        public void fill_dgv(string sql, DataGridView dgv)
        {
            try
            {
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(sql, connectionDB.connection))
                {
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        dgv.DataSource = dt;
                    }
                    else
                    {
                        dgv.Rows.Add("No data");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // to get text from db
        public void resultQuery(string sql, Label result, string column)
        {
            try
            {
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(sql, connectionDB.connection))
                {
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        result.Text = dt.Rows[0][column].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //tampilkan data now di label toolstrip
        public void dateTimeNow(ToolStripLabel dateTimeNow)
        {
            dateTimeNow.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss");
        }

        // for display data in treeview
        public void AddNodes(ref TreeNode node, DataTable dtSource)
        {
            DataTable dt = GetChildData(dtSource, Convert.ToInt32(node.Name));
            foreach (DataRow row in dt.Rows)
            {
                TreeNode childNode = new TreeNode();
                childNode.Name = row["NodeID"].ToString();
                childNode.Text = row["NodeText"].ToString();
                AddNodes(ref childNode, dtSource);
                node.Nodes.Add(childNode);
            }
        }

        public DataTable GetChildData(DataTable dtSource, int parentId)
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[] {
        new DataColumn("NodeId", typeof(int)),
        new DataColumn("ParentId", typeof(int)),
        new DataColumn("NodeText") });
            foreach (DataRow dr in dtSource.Rows)
            {
                if (dr[1].ToString() != parentId.ToString())
                {
                    continue;
                }
                DataRow row = dt.NewRow();
                row["NodeId"] = dr["NodeId"];
                row["ParentId"] = dr["ParentId"];
                row["NodeText"] = dr["NodeText"];
                dt.Rows.Add(row);
            }

            return dt;
        }

        public DataTable GetData(string query)
        {
            using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
            {
                DataTable dt = new DataTable();
                adpt.Fill(dt);
                return dt;
            }
        }

        // Updates all child tree nodes recursively.
        public void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    // If the current node has child nodes, call the CheckAllChildsNodes method recursively.
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        //class for check and uncheck treeview if child checked
        public void SelectParents(TreeNode node, Boolean isChecked)
        {
            var parent = node.Parent;

            if (parent == null)
                return;

            if (!isChecked && HasCheckedNode(parent))
                return;

            parent.Checked = isChecked;
            SelectParents(parent, isChecked);
        }

        public bool HasCheckedNode(TreeNode node)
        {
            return node.Nodes.Cast<TreeNode>().Any(n => n.Checked);
        }

        //uncheck all checkboxes of tree view
        public void UncheckAllNodes(TreeView treeView)
        {
            foreach (TreeNode parent in treeView.Nodes)
            {
                parent.Checked = false;

                foreach (TreeNode child in parent.Nodes)
                {
                    child.Checked = false;
                }
            }
        }

        public void displayCmbList(string sql, string display, string value, ComboBox comboBox)
        {
            try
            {
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(sql, connectionDB.connection))
                {
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        comboBox.DataSource = dt;
                        comboBox.DisplayMember = display;
                        comboBox.ValueMember = value;
                    }
                }
                comboBox.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                connectionDB.connection.Close();
                // tampilkan pesan error
                MessageBox.Show(ex.Message);
            }
        }

        // get selectedcheckedlist
        public void SelectedCheckList(CheckedListBox checkedListBox, string selectedItems)
        {
            selectedItems = string.Empty;

            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                if (checkedListBox.GetItemChecked(i))
                {
                    selectedItems += checkedListBox.GetItemText(checkedListBox.Items[i]).Substring(0, 6) + "\r\n";
                }
            }
        }


        // remove duplicate data in datagridview
        public void RemoveDuplicate(DataGridView grv)
        {
            bool duplicateRow = false;
            if (grv.Rows.Count > 2)
            {
                for (int currentRow = 0; currentRow < grv.Rows.Count - 1; currentRow++)
                {
                    DataGridViewRow rowToCompare = grv.Rows[currentRow];

                    for (int otherRow = currentRow + 1; otherRow < grv.Rows.Count; otherRow++)
                    {
                        DataGridViewRow row = grv.Rows[otherRow];

                        duplicateRow = true;

                        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                        {
                            if (!rowToCompare.Cells[cellIndex].Value.Equals(row.Cells[cellIndex].Value))
                            {
                                duplicateRow = false;
                                break;
                            }
                        }

                        if (duplicateRow)
                        {
                            grv.Rows.Remove(row);
                            otherRow--;
                        }
                    }
                }
            }
        }

        // to display no records data in datagridview c#
        public void norecord_dgv(DataGridView dgv, PaintEventArgs e)
        {
            if (dgv.Rows.Count == 0)
            {
                dgv.ColumnHeadersVisible = false;

                //setting font
                Font font = new Font("Open Sans", 20.0f, FontStyle.Bold);

                //add text no records
                TextRenderer.DrawText(e.Graphics, "No records found.", font,
                dgv.ClientRectangle, Color.White,
                dgv.BackgroundColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else
            {
                dgv.ColumnHeadersVisible = true;
            }
        }

        //get totaldays
        public string totalDay(DateTimePicker picker1, DateTimePicker picker2)
        {
            DateTime startDate = picker1.Value;
            DateTime endDate = picker2.Value;

            TimeSpan ts = endDate - startDate;
            int diffInDay = ts.Days + 1;

            return diffInDay.ToString(); ;
        }

        static int GetWeekNumberOfMonth(DateTime date)
        {
            date = date.Date;
            DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
            DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            if (firstMonthMonday > date)
            {
                firstMonthDay = firstMonthDay.AddMonths(-1);
                firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            }
            return (date - firstMonthMonday).Days / 7 + 1;
        }

        // class for change color 
        public Bitmap MakeGrayscale(Bitmap original)
        {
            Bitmap newBmp = new Bitmap(original.Width, original.Height);
            Graphics g = Graphics.FromImage(newBmp);
            ColorMatrix colorMatrix = new ColorMatrix(
               new float[][]
               {
                   new float[] {.3f, .3f, .3f, 0, 0},
                   new float[] {.59f, .59f, .59f, 0, 0},
                   new float[] {.11f, .11f, .11f, 0, 0},
                   new float[] {0, 0, 0, 1, 0},
                   new float[] {0, 0, 0, 0, 1}
               });
            ImageAttributes img = new ImageAttributes();
            img.SetColorMatrix(colorMatrix);
            g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height), 0, 0, original.Width, original.Height, GraphicsUnit.Pixel, img);
            g.Dispose();
            return newBmp;
        }

        public int TotalNumberOfDaysInMonth(int year, int month)
        {
            return DateTime.DaysInMonth(year, month);
        }
    }
}
