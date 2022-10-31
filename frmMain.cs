using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;
using BIOVMyTimeSheet;
using ClosedXML.Excel;

namespace ReaderEngine
{
    public partial class frmMain : Form
    {        
        private string Sql;
        public frmMain()
        {
            InitializeComponent();
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            ConnectionDB connectionDB = new ConnectionDB();

            try
            {
                var cmd = new MySqlCommand("", connectionDB.connection);
                DateTime time = dateTimePickerTimer.Value;

                string cekdept = "SELECT * FROM tbl_processTimer WHERE time = '" + time.ToString("HH:mm:00") + "'";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(cekdept, connectionDB.connection))
                {
                    DataSet ds = new DataSet();
                    adpt.Fill(ds);

                    // cek jika modelno tsb sudah di upload
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        MessageBox.Show(this, "Unable to add time schedule, Time Schedule already insert", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        connectionDB.connection.Open();
                        string queryAdd = "INSERT INTO tbl_processTimer (time, createDate) VALUES ('" + time.ToString("HH:mm:00") + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                        string[] allQuery = { queryAdd };
                        for (int j = 0; j < allQuery.Length; j++)
                        {
                            cmd.CommandText = allQuery[j];
                            //Masukkan perintah/query yang akan dijalankan ke dalam CommandText
                            cmd.ExecuteNonQuery();
                            //Jalankan perintah / query dalam CommandText pada database
                        }
                        connectionDB.connection.Close();
                        MessageBox.Show(this, "Time Schedule Successfully Added", "Add Time Schedule", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                connectionDB.connection.Close();
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }

        private void refresh()
        {
            // update data in datagridview time result
            dataGridViewProcessTime.DataSource = null;
            dataGridViewProcessTime.Refresh();

            while (dataGridViewProcessTime.Columns.Count > 0)
            {
                dataGridViewProcessTime.Columns.RemoveAt(0);
            }

            loadDataTimer();

            dataGridViewProcessTime.Update();
            dataGridViewProcessTime.Refresh();


            // update data in datagridview transaction result
            dataGridViewTransactionEmployee.DataSource = null;
            dataGridViewTransactionEmployee.Refresh();

            while (dataGridViewTransactionEmployee.Columns.Count > 0)
            {
                dataGridViewTransactionEmployee.Columns.RemoveAt(0);
            }

            loadDataTransaction();

            dataGridViewTransactionEmployee.Update();
            dataGridViewTransactionEmployee.Refresh();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            dateTimeNow.Text = DateTime.Now.ToString("HH:mm");

            loadDataTimer();
            loadDataTransaction();

            ReaderEnginee.BalloonTipIcon = ToolTipIcon.Info;
            ReaderEnginee.BalloonTipTitle = "RFID Reader Data Process";
            ReaderEnginee.BalloonTipText = "Application RFID Reader Data Process";
            ReaderEnginee.ShowBalloonTip(2000);

            timerRefresh.Start();
        }

        private void loadDataTimer()
        {
            ConnectionDB connectionDB = new ConnectionDB();
            try
            {
                string query = "SELECT TIME FROM tbl_processTimer ORDER BY TIME ASC";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
                {
                    DataSet dset = new DataSet();
                    adpt.Fill(dset);
                    dataGridViewProcessTime.DataSource = dset.Tables[0];
                }
            }
            catch (Exception ex)
            {
                connectionDB.connection.Close();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }

        private void loadDataTransaction()
        {
            ConnectionDB connectionDB = new ConnectionDB();
            try
            {
                DateTime dt1 = DateTime.Today.AddDays(-1);
                DateTime dt2 = DateTime.Today;
                dt2 = dt2.AddDays(1).AddSeconds(-1);

                string query = "SELECT e.badgeid, l.rfidno,e.name, e.workarea, l.ipDevice, l.indicator, l.timelog, l.processed FROM tbl_log AS l INNER JOIN tbl_employee AS e " +
                       "ON e.rfidno = l.rfidno WHERE (l.timelog between '" + dt1.ToString("yyyy-MM-dd") + "' and '" + dt2.ToString("yyyy-MM-dd HH:mm:ss") + "') ORDER BY l.timelog, l.id DESC";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
                {
                    DataSet dset = new DataSet();
                    adpt.Fill(dset);
                    dataGridViewTransactionEmployee.DataSource = dset.Tables[0];
                }
            }
            catch (Exception ex)
            {
                connectionDB.connection.Close();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }

        private void processTransaction()
        {
            DateTime dt1 = DateTime.Today.AddDays(-2);

            string sq = "select l.*, e.id as emplid, e.workarea from tbl_log as l inner join tbl_employee as e on e.rfidno = l.rfidno " +
               "where l.timelog>='" + dt1.ToString("yyyy-MM-dd") + "' and processed = 0 order by l.timelog, l.id desc";

            ConnectionDB connectionDB = new ConnectionDB();
            connectionDB.connection.Open();
            using (MySqlDataAdapter da = new MySqlDataAdapter(sq, connectionDB.connection))
            {
                var tmSheet = new Timesheets(connectionDB.connection);
                tmSheet.SetValid2Checkin(12);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    progressBar1.Value = 0;
                    progressBar1.Maximum = dt.Rows.Count;
                    int p = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        p++;
                        progressBar1.Value = p;

                        string msg = "";
                        try
                        {
                            string id = row["rfidno"].ToString();
                            DateTime timeLog = Convert.ToDateTime(row["timelog"]);
                            string sLog = timeLog.ToString("yyyy-MM-dd HH:mm");
                            DateTime dtLog = Convert.ToDateTime(sLog);
                            string flag = row["indicator"].ToString();
                            string lokasi = row["ipdevice"].ToString();

                            tmSheet.ProcessLog(id, dtLog, flag, ref msg);
                            //progressBar1.Value = 0;
                        }
                        catch (Exception ex)
                        {
                            msg = ex.Message;
                        }
                        Application.DoEvents();
                    }
                }
            }

            // export data export
            ExportToExcel();
        }

        private void ExportToExcel()
        {
            ConnectionDB connectionDB = new ConnectionDB();
            try
            {
                string date = DateTime.Now.ToString("dd-MM-yyyy");
                string directoryFile = "\\\\192.168.192.254\\SystemSupport\\Attendance-SMT";
                directoryFile = directoryFile + "\\" + date;
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Late");

                    //to hide gridlines
                    worksheet.ShowGridLines = false;

                    // set column width
                    worksheet.Columns().Width = 15;
                    worksheet.Column(1).Width = 8.43;
                    worksheet.Column(3).Width = 26;

                    worksheet.Rows().Height = 16.25;
                    worksheet.Row(1).Height = 25.5;

                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.25;
                    worksheet.PageSetup.Margins.Left = 0.25;
                    worksheet.PageSetup.Margins.Right = 0;
                    worksheet.PageSetup.Margins.Header = 0.5;
                    worksheet.PageSetup.Margins.Footer = 0.25;
                    worksheet.PageSetup.CenterHorizontally = true;

                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 8)).Merge();
                    worksheet.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheet.Range(worksheet.Cell(2, 8), worksheet.Cell(3, 8)).Style.Font.FontName = "Courier New";
                    worksheet.Range(worksheet.Cell(2, 8), worksheet.Cell(3, 8)).Style.Font.FontSize = 8;
                    worksheet.Range(worksheet.Cell(2, 8), worksheet.Cell(3, 8)).Style.Font.Bold = true;
                    worksheet.Cell(3, 8).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Font.FontName = "Times New Roman";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Font.FontSize = 10;
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(5, 1).Value = "NO";
                    worksheet.Cell(5, 2).Value = "Badge ID";
                    worksheet.Cell(5, 3).Value = "Employee Name";
                    worksheet.Cell(5, 4).Value = "Line Code";
                    worksheet.Cell(5, 5).Value = "Schedule";
                    worksheet.Cell(5, 6).Value = "Actual In";
                    worksheet.Cell(5, 7).Value = "Actual Out";
                    worksheet.Cell(5, 8).Value = "Status";
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Range(worksheet.Cell(5, 1), worksheet.Cell(5, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    worksheet.Cell(5, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    worksheet.Cell(5, 8).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                    int cellRowIndex = 6;
                    int cellColumnIndex = 2;

                    Sql = "SELECT badgeID, NAME, linecode, ScheduleIn, intime, outtime, Sttus FROM (SELECT e.badgeID, e.name, e.linecode, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, DATE_FORMAT(a.outtime, '%H:%i') AS outtime, " +
                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus FROM tbl_attendance a, tbl_employee e WHERE e.id = a.emplid " +
                    "AND a.date = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY intime";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, connectionDB.connection))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex), worksheet.Cell(dt.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                            worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex), worksheet.Cell(dt.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

                            // storing Each row and column value to excel sheet  
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                for (int j = 0; j < dt.Columns.Count; j++)
                                {
                                    worksheet.Cell(i + cellRowIndex, 1).Value = i + 1;
                                    worksheet.Cell(i + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                    if (j == 0)
                                    {
                                        worksheet.Cell(i + cellRowIndex, j + cellColumnIndex).Value = "'" + dt.Rows[i][j].ToString();
                                    }
                                    else
                                    {
                                        worksheet.Cell(i + cellRowIndex, j + cellColumnIndex).Value = dt.Rows[i][j].ToString();
                                    }

                                }
                            }
                            int endPart = dt.Rows.Count + cellRowIndex;

                            // setup border 
                            worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 8)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                            worksheet.Range(worksheet.Cell(cellRowIndex, 8), worksheet.Cell(endPart - 1, 8)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                            worksheet.Range(worksheet.Cell(endPart - 1, 1), worksheet.Cell(endPart - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                            // set value Align center
                            worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    //        var worksheetAbsent = workbook.Worksheets.Add("Absent");

                    //        //to hide gridlines
                    //        worksheetAbsent.ShowGridLines = false;

                    //        // set column width
                    //        worksheetAbsent.Columns().Width = 15;
                    //        worksheetAbsent.Column(1).Width = 8.43;
                    //        worksheetAbsent.Column(3).Width = 26;

                    //        worksheetAbsent.Rows().Height = 16.25;
                    //        worksheetAbsent.Row(1).Height = 25.5;

                    //        worksheetAbsent.PageSetup.Margins.Top = 0.5;
                    //        worksheetAbsent.PageSetup.Margins.Bottom = 0.25;
                    //        worksheetAbsent.PageSetup.Margins.Left = 0.25;
                    //        worksheetAbsent.PageSetup.Margins.Right = 0;
                    //        worksheetAbsent.PageSetup.Margins.Header = 0.5;
                    //        worksheetAbsent.PageSetup.Margins.Footer = 0.25;
                    //        worksheetAbsent.PageSetup.CenterHorizontally = true;

                    //        worksheetAbsent.Range(worksheetAbsent.Cell(1, 1), worksheetAbsent.Cell(1, 8)).Merge();
                    //        worksheetAbsent.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    //        worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    //        worksheetAbsent.Cell(1, 1).Style.Font.FontSize = 20;
                    //        worksheetAbsent.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    //        worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    //        worksheetAbsent.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    //        worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    //        worksheetAbsent.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    //        worksheetAbsent.Range(worksheetAbsent.Cell(2, 8), worksheetAbsent.Cell(3, 8)).Style.Font.FontName = "Courier New";
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(2, 8), worksheetAbsent.Cell(3, 8)).Style.Font.FontSize = 8;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(2, 8), worksheetAbsent.Cell(3, 8)).Style.Font.Bold = true;
                    //        worksheetAbsent.Cell(3, 8).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Font.FontName = "Times New Roman";
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Font.FontSize = 10;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Font.Bold = true;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    //        worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    //        worksheetAbsent.Cell(5, 1).Value = "NO";
                    //        worksheetAbsent.Cell(5, 2).Value = "Badge ID";
                    //        worksheetAbsent.Cell(5, 3).Value = "Employee Name";
                    //        worksheetAbsent.Cell(5, 4).Value = "Line Code";
                    //        worksheetAbsent.Cell(5, 5).Value = "Schedule";
                    //        worksheetAbsent.Cell(5, 6).Value = "Actual In";
                    //        worksheetAbsent.Cell(5, 7).Value = "Actual Out";
                    //        worksheetAbsent.Cell(5, 8).Value = "Status";
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    //        worksheetAbsent.Range(worksheetAbsent.Cell(5, 1), worksheetAbsent.Cell(5, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    //        worksheetAbsent.Cell(5, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    //        worksheetAbsent.Cell(5, 8).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                    //        int cellRowIndexworksheetAbsent = 6;
                    //        int cellColumnIndexworksheetAbsent = 2;

                    //        Sql = "(SELECT badgeID, NAME, linecode, ScheduleIn, intime, outtime, Sttus FROM (SELECT badgeID, NAME, linecode, '-' AS ScheduleIn, " +
                    //"'-' AS intime, '-' AS outtime, 'Absent' AS Sttus FROM tbl_employee WHERE badgeID NOT IN (SELECT b.badgeID FROM tbl_attendance a, tbl_employee b " +
                    //"WHERE a.EmplId = b.id AND a.date = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' AND intime IS NOT NULL) ) AS A ) ORDER BY intime";

                    //        using (MySqlDataAdapter adptAbsent = new MySqlDataAdapter(Sql, connectionDB.connection))
                    //        {
                    //            DataTable dtAbsent = new DataTable();
                    //            adptAbsent.Fill(dtAbsent);

                    //            if (dtAbsent.Rows.Count > 0)
                    //            {
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, cellColumnIndexworksheetAbsent), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndexworksheetAbsent, 9)).Style.Font.FontName = "Times New Roman";
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, cellColumnIndexworksheetAbsent), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndexworksheetAbsent, 9)).Style.Font.FontSize = 9;

                    //                // storing Each row and column value to excel sheet  
                    //                for (int i = 0; i < dtAbsent.Rows.Count; i++)
                    //                {
                    //                    for (int j = 0; j < dtAbsent.Columns.Count; j++)
                    //                    {
                    //                        worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, 1).Value = i + 1;
                    //                        worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    //                        if (j == 0)
                    //                        {
                    //                            worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, j + cellColumnIndexworksheetAbsent).Value = "'" + dtAbsent.Rows[i][j].ToString();
                    //                        }
                    //                        else
                    //                        {
                    //                            worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, j + cellColumnIndexworksheetAbsent).Value = dtAbsent.Rows[i][j].ToString();
                    //                        }

                    //                    }
                    //                }
                    //                int endPartAbsent = dtAbsent.Rows.Count + cellRowIndexworksheetAbsent;

                    //                // setup border 
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 1), worksheetAbsent.Cell(endPartAbsent - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 8)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 1), worksheetAbsent.Cell(endPartAbsent - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 8), worksheetAbsent.Cell(endPartAbsent - 1, 8)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(endPartAbsent - 1, 1), worksheetAbsent.Cell(endPartAbsent - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                    //                // set value Align center
                    //                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    //            }
                    //        }

                            workbook.SaveAs(directoryFile + "\\Summary.xlsx");
                        }
                        else
                        {
                            workbook.SaveAs(directoryFile + "\\Summary.xlsx");
                        }
                    }
                }                
                //MessageBox.Show(this, "Excel File Success Generated", "Generate Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // tampilkan pesan error
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }

        private void dataGridViewProcessTime_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i;
            i = dataGridViewProcessTime.SelectedCells[0].RowIndex;
            string timeslctd = dataGridViewProcessTime.Rows[i].Cells[0].Value.ToString();
            if (e.ColumnIndex == 0)
            {
                dateTimePickerTimer.Text = timeslctd;
            }
        }

        private void deleteBtn_Click(object sender, EventArgs e)
        {
            ConnectionDB connectionDB = new ConnectionDB();
            int i;
            i = dataGridViewProcessTime.SelectedCells[0].RowIndex;
            string timeslctd = dataGridViewProcessTime.Rows[i].Cells[0].Value.ToString();

            string message = "Do you want to delete this Timer " + timeslctd + "?";
            string title = "Delete Timer";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Information;
            DialogResult result = MessageBox.Show(this, message, title, buttons, icon);

            if (result == DialogResult.Yes)
            {
                try
                {
                    var cmd = new MySqlCommand("", connectionDB.connection);

                    connectionDB.connection.Open();
                    string querydelete = "DELETE FROM tbl_processTimer WHERE TIME = '" + timeslctd + "'";
                    cmd.CommandText = querydelete;
                    cmd.ExecuteNonQuery();
                    connectionDB.connection.Close();
                    MessageBox.Show("Record Deleted successfully", "Timer List Record", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    refresh();
                }
                catch (Exception ex)
                {
                    connectionDB.connection.Close();
                    MessageBox.Show("Unable to remove selected Timer", "Timer List Record", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //MessageBox.Show(ex.Message);
                }

                finally
                {
                    connectionDB.connection.Close();
                }
            }
        }

        private void timerRefresh_Tick(object sender, EventArgs e)
        {
            refresh();
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            ReaderEnginee.BalloonTipTitle = "RFID Reader Process";
            ReaderEnginee.BalloonTipText = "Application RFID Reader Process";

            if (FormWindowState.Minimized == this.WindowState)
            {
                this.ShowInTaskbar = false;
                ReaderEnginee.Visible = true;
                ReaderEnginee.ShowBalloonTip(4);
                this.Hide();
            }
            else if (FormWindowState.Normal == this.WindowState)
            {
                ReaderEnginee.Visible = false;
            }
        }

        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            ConnectionDB connectionDB = new ConnectionDB();
            string now = null;

            dateTimeNow.Text = DateTime.Now.ToString("HH:mm");
            now = DateTime.Now.ToString("HH:mm:00");

            // running process attendance
            try
            {
                string query = "SELECT TIME FROM tbl_processTimer WHERE TIME = '" + now + "'";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
                {
                    DataSet dset = new DataSet();
                    adpt.Fill(dset);
                    if (dset.Tables[0].Rows.Count > 0)
                    {
                        processTransaction();
                    }
                }
            }
            catch (Exception ex)
            {
                connectionDB.connection.Close();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }
    }
}
