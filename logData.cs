using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;
using BIOVMyTimeSheet;
using ClosedXML.Excel;

namespace ReaderEngine
{
    public partial class logData : Form
    {
        private string Sql;
        MySqlConnection myConn;
        string fileReport;

        string year;
        string month;
        string dateReport;

        string startDate;
        string endDate;

        readonly Helper help = new Helper();

        public logData()
        {
            InitializeComponent();
        }

        private void logData_Load(object sender, EventArgs e)
        {
            loadDataTransaction();
        }

        // function to reset data late employee in month
        private void resetLate()
        {
            string koneksi = ConnectionDB.strProvider;
            myConn = new MySqlConnection(koneksi);
            try
            {
                var cmd = new MySqlCommand("", myConn);
                myConn.Open();
                string queryReset = "UPDATE tbl_attendance SET isLate = '0' WHERE DATE >= '" + startDate + "' AND DATE <= '" + endDate + "'";
                cmd.CommandText = queryReset;
                cmd.ExecuteNonQuery();

                string queryUpdate = "UPDATE tbl_attendance SET isLate = '1' WHERE id IN (SELECT id FROM " +
                    "(SELECT b.name, a.id, a.date, IF(intime > ScheduleIn, 'Late', 'Ontime') AS Sttus FROM tbl_attendance a, tbl_employee b " +
                    "WHERE  b.id = a.emplid AND a.date >= '" + startDate + "' AND a.date <= '" + endDate + "') AS a WHERE sttus = 'late')";
                cmd.CommandText = queryUpdate;
                cmd.ExecuteNonQuery();

                myConn.Close();
            }
            catch (System.Exception ex)
            {
                myConn.Close();
            }
            finally
            {
                myConn.Dispose();
            }
        }


        private void refresh()
        {
            // update data in datagridview transaction result
            this.dataGridViewTransactionEmployee.DoubleBuffered(true);
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

        private void loadDataTransaction()
        {
            ConnectionDB connectionDB = new ConnectionDB();
            try
            {
                DateTime dt1 = dateTimePickerDate.Value.AddDays(-1);
                DateTime dt2 = dateTimePickerDate.Value.AddDays(1).AddSeconds(-1);

                string query = "SELECT e.badgeid, l.rfidno,e.name, e.workarea, l.ipDevice, l.indicator, l.timelog, l.processed FROM tbl_log AS l INNER JOIN tbl_employee AS e " +
                       "ON e.rfidno = l.rfidno WHERE (l.timelog between '" + dt1.ToString("yyyy-MM-dd") + "' and '" + dt2.ToString("yyyy-MM-dd HH:mm:ss") + "') ORDER BY l.timelog, l.id DESC";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
                {
                    DataSet dset = new DataSet();
                    adpt.Fill(dset);
                    dataGridViewTransactionEmployee.DataSource = dset.Tables[0];
                }
            }
            catch (System.Exception ex)
            {
                connectionDB.connection.Close();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connectionDB.connection.Close();
            }
        }

        // function to reset data employee attandance depend on selected date
        private void resetAttendanceData()
        {
            string koneksi = ConnectionDB.strProvider;
            myConn = new MySqlConnection(koneksi);
            try
            {
                var cmd = new MySqlCommand("", myConn);

                // reset attendance employee
                myConn.Open();
                string queryUpdate = " UPDATE tbl_attendance SET shiftId = Null, ScheduleIn = Null, ScheduleOut = Null, intime = Null, outtime = Null  " +
                    ", BreakOut = Null, BreakIn = Null, TotalBreak = 0, LateIn = 0, EarlyOut = 0 WHERE date = '" + dateTimePickerDate.Value.ToString("yyyy-MM-dd") + "'";

                cmd.CommandText = queryUpdate;
                cmd.ExecuteNonQuery();

                myConn.Close();
            }
            catch (System.Exception ex)
            {
                myConn.Close();
            }
            finally
            {
                myConn.Dispose();
            }            
        }

        // function to reset flag processed in tbl_log depent on date selected
        private void resetFlagLog()
        {
            string koneksi = ConnectionDB.strProvider;
            myConn = new MySqlConnection(koneksi);
            try
            {
                var cmd = new MySqlCommand("", myConn);

                // reset attendance employee
                myConn.Open();
                string queryUpdate = "UPDATE tbl_log SET processed = 0 WHERE timelog > '" + dateTimePickerDate.Value.ToString("yyyy-MM-dd") + " 00:00:00' AND timelog < '" + dateTimePickerDate.Value.AddDays(1).ToString("yyyy-MM-dd") + " 00:00:00'";


                cmd.CommandText = queryUpdate;
                cmd.ExecuteNonQuery();

                myConn.Close();
            }
            catch (System.Exception ex)
            {
                myConn.Close();
            }
            finally
            {
                myConn.Dispose();
            }            
        }

        private void processTransaction()
        {
            string koneksi = ConnectionDB.strProvider;
            myConn = new MySqlConnection(koneksi);
            try
            {
                myConn.Open();

                DateTime dt1 = dateTimePickerDate.Value.AddDays(-1);
                DateTime dt2 = dateTimePickerDate.Value.AddDays(1);

                string sq = "select l.*, e.id as emplid, e.workarea from tbl_log as l inner join tbl_employee as e on e.rfidno = l.rfidno " +
                   "where l.timelog>='" + dt1.ToString("yyyy-MM-dd") + "'  and l.timelog<='" + dt2.ToString("yyyy-MM-dd") + "'  and processed = 0 order by l.timelog, l.id desc";

                using (MySqlDataAdapter da = new MySqlDataAdapter(sq, myConn))
                {
                    var tmSheet = new Timesheets(myConn);
                    tmSheet.SetValid2Checkin(15);
                    tmSheet.SetValidBreakSeconds(60);

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
                            catch (System.Exception ex)
                            {
                                msg = ex.Message;
                            }
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
                myConn.Dispose();
            }
        }

        private void ExportToExcel()
        {
            try
            {
                string koneksi = ConnectionDB.strProvider;
                myConn = new MySqlConnection(koneksi);

                string date = dateTimePickerDate.Value.ToString("yyyy-MM-dd");

                string directoryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                directoryFile = directoryFile + "\\Attendance-SMT";
                using (var workbook = new XLWorkbook())
                {
                    // late sheet
                    var worksheet = workbook.Worksheets.Add("Late");

                    // reset data late employee
                    resetLate();

                    //to hide gridlines
                    worksheet.ShowGridLines = false;

                    // set column width
                    worksheet.Columns().Width = 15;
                    worksheet.Column(1).Width = 5;
                    worksheet.Column(2).Width = 14;
                    worksheet.Column(3).Width = 31;
                    worksheet.Column(4).Width = 9;
                    worksheet.Column(5).Width = 9;
                    worksheet.Column(6).Width = 9;
                    worksheet.Column(7).Width = 9;
                    worksheet.Column(8).Width = 17;
                    worksheet.Column(9).Width = 23;
                    worksheet.Column(10).Width = 23;


                    worksheet.Rows().Height = 16.25;
                    worksheet.Row(1).Height = 25.5;

                    // set format hour
                    worksheet.Column(6).Style.NumberFormat.Format = "hh:mm";
                    worksheet.Column(7).Style.NumberFormat.Format = "hh:mm";

                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.25;
                    worksheet.PageSetup.Margins.Left = 0.25;
                    worksheet.PageSetup.Margins.Right = 0;
                    worksheet.PageSetup.Margins.Header = 0.5;
                    worksheet.PageSetup.Margins.Footer = 0.25;
                    worksheet.PageSetup.CenterHorizontally = true;

                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Merge();
                    worksheet.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.FontName = "Times New Roman";
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.FontSize = 9;
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(2, 10), worksheet.Cell(3, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheet.Range(worksheet.Cell(2, 10), worksheet.Cell(3, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheet.Cell(2, 1).Value = "Attendance Marked At";
                    worksheet.Cell(2, 3).Value = ": " + date;
                    worksheet.Cell(2, 9).Value = "Report Date:";
                    worksheet.Cell(2, 10).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    int cellRowIndexStartLate = 3;
                    int cellRowIndexlate = 0;
                    int totalLate = 0;

                    // find workarea
                    Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, " +
                        "DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, DATE_FORMAT(a.intime, '%H:%i') AS intime, " +
                        "TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus " +
                        "FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f WHERE e.id = a.emplid AND e.linecode = f.name AND " +
                        "a.date = '" + date + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' " +
                        "GROUP BY workarea ORDER BY FIELD(workarea, 'SMT', 'SMT-DIPPING', 'SMT-SA', 'SMT-MAINROOM') ";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                string workarea = dt.Rows[i][0].ToString();
                                int total = Convert.ToInt32(dt.Rows[i][1].ToString());
                                total += total;

                                // set header excel
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.FontName = "Times New Roman";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 10)).Style.Font.FontSize = 9;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.FontSize = 10;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.Bold = true;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1).Value = "Workarea";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 3).Value = ": " + workarea;
                                //worksheet.Cell(3, 9).Value = "Total Late :" + total;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Value = "NO";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 2).Value = "Badge ID";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 3).Value = "Employee Name";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 4).Value = "Line Code";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 5).Value = "Section";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 6).Value = "Schedule";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 7).Value = "Actual In";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 8).Value = "Total Late (Minute)";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9).Value = "Total Late Days In a Month";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10).Value = "Reason";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                int cellRowIndex = cellRowIndexStartLate + cellRowIndexlate + 2;
                                int cellColumnIndex = 2;

                                Sql = "SELECT t1.badgeID, t1.NAME, t1.linecode, t1.DESCRIPTION, t1.ScheduleIn, t1.intime, t1.diff, t2.totallate, t1.reason FROM " +
                                    "(SELECT badgeID, NAME, linecode, DESCRIPTION, ScheduleIn, intime, diff, reason FROM " +
                                    "(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                    "WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '" + workarea + "' AND  a.ScheduleIn IS NOT NULL " +
                                    "ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME )t1 " +
                                    "LEFT JOIN (SELECT b.badgeID, SUM(a.islate) AS totalLate FROM tbl_attendance a, tbl_employee b, tbl_masterlinecode c " +
                                    "WHERE a.emplid = b.id AND c.name = b.linecode AND b.workarea = '" + workarea + "' AND(a.date >= '" + startDate + "' AND a.date <= '" + endDate + "') " +
                                    "AND b.badgeID IN(SELECT badgeID FROM " +
                                    "(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff," +
                                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                    "WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '" + workarea + "' AND  a.ScheduleIn IS NOT NULL " +
                                    "ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late') GROUP BY b.badgeID) t2 ON  t1.badgeID = t2.badgeID";

                                //"SELECT badgeID, NAME, linecode, DESCRIPTION, ScheduleIn, intime, diff, reason FROM " +
                                //"(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                //"DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                                //"IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                //"WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '"+workarea+"' AND  a.ScheduleIn IS NOT NULL " +
                                //"ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME";

                                using (MySqlDataAdapter adptLateWorkarea = new MySqlDataAdapter(Sql, myConn))
                                {
                                    DataTable dtLateWorkarea = new DataTable();
                                    adptLateWorkarea.Fill(dtLateWorkarea);

                                    if (dtLateWorkarea.Rows.Count > 0)
                                    {
                                        totalLate = totalLate + dtLateWorkarea.Rows.Count;
                                        worksheet.Cell(3, 9).Value = "Total Late :";
                                        worksheet.Cell(3, 10).Value = totalLate;

                                        worksheet.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 10)).Style.Font.FontName = "Times New Roman";
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 10)).Style.Font.FontSize = 9;

                                        // storing Each row and column value to excel sheet  
                                        for (int x = 0; x < dtLateWorkarea.Rows.Count; x++)
                                        {
                                            for (int y = 0; y < dtLateWorkarea.Columns.Count; y++)
                                            {
                                                worksheet.Cell(x + cellRowIndex, 1).Value = x + 1;
                                                worksheet.Cell(x + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                                if (y == 0)
                                                {
                                                    worksheet.Cell(x + cellRowIndex, y + cellColumnIndex).Value = "'" + dtLateWorkarea.Rows[x][y].ToString();
                                                }
                                                else
                                                {
                                                    worksheet.Cell(x + cellRowIndex, y + cellColumnIndex).Value = dtLateWorkarea.Rows[x][y].ToString();
                                                }

                                                if (Convert.ToInt32(dtLateWorkarea.Rows[x][6].ToString()) > 31)
                                                {
                                                    worksheet.Cell(x + cellRowIndex, 8).Style.Fill.BackgroundColor = XLColor.Yellow;
                                                }
                                            }
                                        }
                                        int endPart = dtLateWorkarea.Rows.Count + cellRowIndex;

                                        // setup border 
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 10)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 10), worksheet.Cell(endPart - 1, 10)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(endPart - 1, 1), worksheet.Cell(endPart - 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                                        // set value Align center
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                        cellRowIndexlate = endPart - 2;
                                    }
                                }
                            }

                            worksheet.Cell(cellRowIndexlate + 3, 1).Value = "*Note : Employee with yellow mark probably missing scan-in, please make sure employee check name in dashboard before entry work area";
                        }
                        else
                        {
                            worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 10)).Merge();
                            worksheet.Cell(4, 1).Style.Font.FontName = "Times New Roman";
                            worksheet.Cell(4, 1).Style.Font.Bold = true;
                            worksheet.Cell(4, 1).Style.Font.FontSize = 12;
                            worksheet.Cell(4, 1).Style.Font.FontColor = XLColor.Black;
                            worksheet.Cell(4, 1).Style.Font.Bold = true;
                            worksheet.Cell(4, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheet.Cell(4, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheet.Cell(4, 1).Value = "No Any Data";
                            worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }

                    // sheet Absent     
                    var worksheetAbsent = workbook.Worksheets.Add("Absent");
                    //to hide gridlines
                    worksheetAbsent.ShowGridLines = false;

                    // set column width
                    worksheetAbsent.Columns().Width = 15;
                    worksheetAbsent.Column(1).Width = 5;
                    worksheetAbsent.Column(2).Width = 14;
                    worksheetAbsent.Column(3).Width = 31;

                    worksheetAbsent.Rows().Height = 16.25;
                    worksheetAbsent.Row(1).Height = 25.5;

                    worksheetAbsent.PageSetup.Margins.Top = 0.5;
                    worksheetAbsent.PageSetup.Margins.Bottom = 0.25;
                    worksheetAbsent.PageSetup.Margins.Left = 0.25;
                    worksheetAbsent.PageSetup.Margins.Right = 0;
                    worksheetAbsent.PageSetup.Margins.Header = 0.5;
                    worksheetAbsent.PageSetup.Margins.Footer = 0.25;
                    worksheetAbsent.PageSetup.CenterHorizontally = true;

                    worksheetAbsent.Range(worksheetAbsent.Cell(1, 1), worksheetAbsent.Cell(1, 7)).Merge();
                    worksheetAbsent.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    worksheetAbsent.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheetAbsent.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    worksheetAbsent.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheetAbsent.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.FontName = "Times New Roman";
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.FontSize = 9;
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.Bold = true;
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 4), worksheetAbsent.Cell(3, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheetAbsent.Cell(2, 1).Value = "Attendance Marked At";
                    worksheetAbsent.Cell(2, 3).Value = ": " + date;
                    worksheetAbsent.Cell(2, 6).Value = "Report Date:";
                    worksheetAbsent.Cell(2, 7).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    int cellRowIndexStartAbsent = 3;
                    int cellRowIndexAbsent = 0;
                    int totalAbsent = 0;

                    // find workarea
                    Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT badgeID, NAME, linecode, DESCRIPTION, workarea FROM (SELECT a.badgeID, a.NAME, a.linecode, b.description, a.workarea " +
                        "FROM tbl_employee a, tbl_masterlinecode b WHERE a.linecode = b.name AND a.status = 1 AND badgeID NOT IN(SELECT b.badgeID FROM tbl_attendance a, tbl_employee b " +
                        "WHERE a.EmplId = b.id AND a.date = '" + date + "' AND a.intime IS NOT NULL)) AS A ) AS A GROUP BY workarea ORDER BY FIELD(workarea, 'SMT', 'SMT-DIPPING', 'SMT-SA', 'SMT-MAINROOM')";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                string workarea = dt.Rows[i][0].ToString();
                                int total = Convert.ToInt32(dt.Rows[i][1].ToString());
                                total += total;

                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.FontName = "Times New Roman";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.FontSize = 10;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.Bold = true;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1).Value = "Workarea";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3).Value = ": " + workarea;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.FontName = "Times New Roman";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.FontSize = 9;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.Bold = true;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1).Value = "NO";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 2).Value = "Badge ID";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 3).Value = "Employee Name";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 4).Value = "Line Code";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 5).Value = "Section";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 6).Value = "Work Area";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7).Value = "Reason";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                int cellRowIndex = cellRowIndexStartAbsent + cellRowIndexAbsent + 2;
                                int cellColumnIndex = 2;

                                Sql = "(SELECT badgeID, NAME, linecode, DESCRIPTION, workarea, reason FROM (SELECT a.badgeID, a.NAME, a.linecode, b.description, a.workarea, c.reason " +
                        "FROM tbl_employee a, tbl_masterlinecode b, tbl_attendance c WHERE a.linecode = b.name AND a.id = c.emplID AND a.status = 1 AND c.date = '" + date + "'  AND badgeID " +
                        "NOT IN(SELECT b.badgeID FROM tbl_attendance a, tbl_employee b WHERE a.EmplId = b.id AND a.date = '" + date + "' AND a.intime IS NOT NULL)) AS A " +
                        "WHERE workarea = '" + workarea + "' ) ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME";

                                using (MySqlDataAdapter adptAbsent = new MySqlDataAdapter(Sql, myConn))
                                {
                                    DataTable dtAbsent = new DataTable();
                                    adptAbsent.Fill(dtAbsent);

                                    if (dtAbsent.Rows.Count > 0)
                                    {
                                        totalAbsent = totalAbsent + dtAbsent.Rows.Count;

                                        worksheetAbsent.Cell(3, 6).Value = "Total Absent :";
                                        worksheetAbsent.Cell(3, 7).Value = totalAbsent;
                                        worksheetAbsent.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, cellColumnIndex - 1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, cellColumnIndex - 1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

                                        // storing Each row and column value to excel sheet  
                                        for (int x = 0; x < dtAbsent.Rows.Count; x++)
                                        {
                                            for (int y = 0; y < dtAbsent.Columns.Count; y++)
                                            {
                                                worksheetAbsent.Cell(x + cellRowIndex, 1).Value = x + 1;
                                                worksheetAbsent.Cell(x + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                                if (y == 0)
                                                {
                                                    worksheetAbsent.Cell(x + cellRowIndex, y + cellColumnIndex).Value = "'" + dtAbsent.Rows[x][y].ToString();
                                                }
                                                else
                                                {
                                                    worksheetAbsent.Cell(x + cellRowIndex, y + cellColumnIndex).Value = dtAbsent.Rows[x][y].ToString();
                                                }
                                            }
                                        }
                                        int endPartAbsent = dtAbsent.Rows.Count + cellRowIndex;

                                        // setup border 
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 1), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 1), worksheetAbsent.Cell(endPartAbsent - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 7), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(endPartAbsent - 1, 1), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                                        // set value Align center
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                        cellRowIndexAbsent = endPartAbsent - 2;
                                    }
                                }
                            }
                            //worksheetAbsent.Cell(cellRowIndexAbsent + 3, 1).Value = "*Note : Break more than 90 Minutes";
                        }
                        else
                        {
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 7)).Merge();
                            worksheetAbsent.Cell(4, 1).Style.Font.FontName = "Times New Roman";
                            worksheetAbsent.Cell(4, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(4, 1).Style.Font.FontSize = 12;
                            worksheetAbsent.Cell(4, 1).Style.Font.FontColor = XLColor.Black;
                            worksheetAbsent.Cell(4, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(4, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheetAbsent.Cell(4, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheetAbsent.Cell(4, 1).Value = "No Any Data";
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }
                    workbook.SaveAs(directoryFile + "\\" + year + "\\Summary " + date + ".xlsx");
                }
                fileReport = directoryFile + "\\" + year + "\\Summary " + date + ".xlsx";
                System.Diagnostics.Process.Start(@"" + directoryFile + "\\" + year + "\\Summary " + date + ".xlsx");
                MessageBox.Show(this, "Excel File Success Generated", "Generate Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // tampilkan pesan error
                MessageBox.Show(ex.Message);
            }
        }

        private void ExportToExcelEndMonth()
        {
            try
            {
                string koneksi = ConnectionDB.strProvider;
                myConn = new MySqlConnection(koneksi);

                string date = dateTimePickerDate.Value.ToString("yyyy-MM-dd");

                string directoryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                directoryFile = directoryFile + "\\Attendance-SMT";
                using (var workbook = new XLWorkbook())
                {
                    // late sheet
                    var worksheet = workbook.Worksheets.Add("Late");

                    // reset data late employee
                    resetLate();

                    //to hide gridlines
                    worksheet.ShowGridLines = false;

                    // set column width
                    worksheet.Columns().Width = 15;
                    worksheet.Column(1).Width = 5;
                    worksheet.Column(2).Width = 14;
                    worksheet.Column(3).Width = 31;
                    worksheet.Column(4).Width = 9;
                    worksheet.Column(5).Width = 9;
                    worksheet.Column(6).Width = 9;
                    worksheet.Column(7).Width = 9;
                    worksheet.Column(8).Width = 17;
                    worksheet.Column(9).Width = 23;
                    worksheet.Column(10).Width = 23;


                    worksheet.Rows().Height = 16.25;
                    worksheet.Row(1).Height = 25.5;

                    // set format hour
                    worksheet.Column(6).Style.NumberFormat.Format = "hh:mm";
                    worksheet.Column(7).Style.NumberFormat.Format = "hh:mm";

                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.25;
                    worksheet.PageSetup.Margins.Left = 0.25;
                    worksheet.PageSetup.Margins.Right = 0;
                    worksheet.PageSetup.Margins.Header = 0.5;
                    worksheet.PageSetup.Margins.Footer = 0.25;
                    worksheet.PageSetup.CenterHorizontally = true;

                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Merge();
                    worksheet.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.FontName = "Times New Roman";
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.FontSize = 9;
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 10)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(2, 10), worksheet.Cell(3, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheet.Range(worksheet.Cell(2, 10), worksheet.Cell(3, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheet.Cell(2, 1).Value = "Attendance Marked At";
                    worksheet.Cell(2, 3).Value = ": " + date;
                    worksheet.Cell(2, 9).Value = "Report Date:";
                    worksheet.Cell(2, 10).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    int cellRowIndexStartLate = 3;
                    int cellRowIndexlate = 0;
                    int totalLate = 0;

                    // find workarea
                    Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, " +
                        "DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, DATE_FORMAT(a.intime, '%H:%i') AS intime, " +
                        "TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus " +
                        "FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f WHERE e.id = a.emplid AND e.linecode = f.name AND " +
                        "a.date = '" + date + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' " +
                        "GROUP BY workarea ORDER BY FIELD(workarea, 'SMT', 'SMT-DIPPING', 'SMT-SA', 'SMT-MAINROOM') ";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                string workarea = dt.Rows[i][0].ToString();
                                int total = Convert.ToInt32(dt.Rows[i][1].ToString());
                                total += total;

                                // set header excel
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.FontName = "Times New Roman";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 10)).Style.Font.FontSize = 9;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.FontSize = 10;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Font.Bold = true;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1).Value = "Workarea";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 3).Value = ": " + workarea;
                                //worksheet.Cell(3, 9).Value = "Total Late :" + total;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Value = "NO";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 2).Value = "Badge ID";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 3).Value = "Employee Name";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 4).Value = "Line Code";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 5).Value = "Section";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 6).Value = "Schedule";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 7).Value = "Actual In";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 8).Value = "Total Late (Minute)";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9).Value = "Total Late Days In a Month";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10).Value = "Reason";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 10).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                int cellRowIndex = cellRowIndexStartLate + cellRowIndexlate + 2;
                                int cellColumnIndex = 2;

                                Sql = "SELECT t1.badgeID, t1.NAME, t1.linecode, t1.DESCRIPTION, t1.ScheduleIn, t1.intime, t1.diff, t2.totallate, t1.reason FROM " +
                                    "(SELECT badgeID, NAME, linecode, DESCRIPTION, ScheduleIn, intime, diff, reason FROM " +
                                    "(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                    "WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '" + workarea + "' AND  a.ScheduleIn IS NOT NULL " +
                                    "ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME )t1 " +
                                    "LEFT JOIN (SELECT b.badgeID, SUM(a.islate) AS totalLate FROM tbl_attendance a, tbl_employee b, tbl_masterlinecode c " +
                                    "WHERE a.emplid = b.id AND c.name = b.linecode AND b.workarea = '" + workarea + "' AND(a.date >= '" + startDate + "' AND a.date <= '" + endDate + "') " +
                                    "AND b.badgeID IN(SELECT badgeID FROM " +
                                    "(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff," +
                                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                    "WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '" + workarea + "' AND  a.ScheduleIn IS NOT NULL " +
                                    "ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late') GROUP BY b.badgeID) t2 ON  t1.badgeID = t2.badgeID";

                                //"SELECT badgeID, NAME, linecode, DESCRIPTION, ScheduleIn, intime, diff, reason FROM " +
                                //"(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                //"DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                                //"IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus, a.reason FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f " +
                                //"WHERE e.id = a.emplid AND e.linecode = f.name AND a.date = '" + date + "' AND e.workarea = '"+workarea+"' AND  a.ScheduleIn IS NOT NULL " +
                                //"ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME";

                                using (MySqlDataAdapter adptLateWorkarea = new MySqlDataAdapter(Sql, myConn))
                                {
                                    DataTable dtLateWorkarea = new DataTable();
                                    adptLateWorkarea.Fill(dtLateWorkarea);

                                    if (dtLateWorkarea.Rows.Count > 0)
                                    {
                                        totalLate = totalLate + dtLateWorkarea.Rows.Count;
                                        worksheet.Cell(3, 9).Value = "Total Late :";
                                        worksheet.Cell(3, 10).Value = totalLate;

                                        worksheet.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 10)).Style.Font.FontName = "Times New Roman";
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 10)).Style.Font.FontSize = 9;

                                        // storing Each row and column value to excel sheet  
                                        for (int x = 0; x < dtLateWorkarea.Rows.Count; x++)
                                        {
                                            for (int y = 0; y < dtLateWorkarea.Columns.Count; y++)
                                            {
                                                worksheet.Cell(x + cellRowIndex, 1).Value = x + 1;
                                                worksheet.Cell(x + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                                if (y == 0)
                                                {
                                                    worksheet.Cell(x + cellRowIndex, y + cellColumnIndex).Value = "'" + dtLateWorkarea.Rows[x][y].ToString();
                                                }
                                                else
                                                {
                                                    worksheet.Cell(x + cellRowIndex, y + cellColumnIndex).Value = dtLateWorkarea.Rows[x][y].ToString();
                                                }

                                                if (Convert.ToInt32(dtLateWorkarea.Rows[x][6].ToString()) > 31)
                                                {
                                                    worksheet.Cell(x + cellRowIndex, 8).Style.Fill.BackgroundColor = XLColor.Yellow;
                                                }
                                            }
                                        }
                                        int endPart = dtLateWorkarea.Rows.Count + cellRowIndex;

                                        // setup border 
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 10)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 10), worksheet.Cell(endPart - 1, 10)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(endPart - 1, 1), worksheet.Cell(endPart - 1, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                                        // set value Align center
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                        cellRowIndexlate = endPart - 2;
                                    }
                                }
                            }

                            worksheet.Cell(cellRowIndexlate + 3, 1).Value = "*Note : Employee with yellow mark probably missing scan-in, please make sure employee check name in dashboard before entry work area";
                        }
                        else
                        {
                            worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 10)).Merge();
                            worksheet.Cell(4, 1).Style.Font.FontName = "Times New Roman";
                            worksheet.Cell(4, 1).Style.Font.Bold = true;
                            worksheet.Cell(4, 1).Style.Font.FontSize = 12;
                            worksheet.Cell(4, 1).Style.Font.FontColor = XLColor.Black;
                            worksheet.Cell(4, 1).Style.Font.Bold = true;
                            worksheet.Cell(4, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheet.Cell(4, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheet.Cell(4, 1).Value = "No Any Data";
                            worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }

                    // sheet Absent     
                    var worksheetAbsent = workbook.Worksheets.Add("Absent");
                    //to hide gridlines
                    worksheetAbsent.ShowGridLines = false;

                    // set column width
                    worksheetAbsent.Columns().Width = 15;
                    worksheetAbsent.Column(1).Width = 5;
                    worksheetAbsent.Column(2).Width = 14;
                    worksheetAbsent.Column(3).Width = 31;

                    worksheetAbsent.Rows().Height = 16.25;
                    worksheetAbsent.Row(1).Height = 25.5;

                    worksheetAbsent.PageSetup.Margins.Top = 0.5;
                    worksheetAbsent.PageSetup.Margins.Bottom = 0.25;
                    worksheetAbsent.PageSetup.Margins.Left = 0.25;
                    worksheetAbsent.PageSetup.Margins.Right = 0;
                    worksheetAbsent.PageSetup.Margins.Header = 0.5;
                    worksheetAbsent.PageSetup.Margins.Footer = 0.25;
                    worksheetAbsent.PageSetup.CenterHorizontally = true;

                    worksheetAbsent.Range(worksheetAbsent.Cell(1, 1), worksheetAbsent.Cell(1, 7)).Merge();
                    worksheetAbsent.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    worksheetAbsent.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheetAbsent.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                    worksheetAbsent.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheetAbsent.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.FontName = "Times New Roman";
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.FontSize = 9;
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 7)).Style.Font.Bold = true;
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 4), worksheetAbsent.Cell(3, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheetAbsent.Cell(2, 1).Value = "Attendance Marked At";
                    worksheetAbsent.Cell(2, 3).Value = ": " + date;
                    worksheetAbsent.Cell(2, 6).Value = "Report Date:";
                    worksheetAbsent.Cell(2, 7).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    int cellRowIndexStartAbsent = 3;
                    int cellRowIndexAbsent = 0;
                    int totalAbsent = 0;

                    // find workarea
                    Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT badgeID, NAME, linecode, DESCRIPTION, workarea FROM (SELECT a.badgeID, a.NAME, a.linecode, b.description, a.workarea " +
                        "FROM tbl_employee a, tbl_masterlinecode b WHERE a.linecode = b.name AND a.status = 1 AND badgeID NOT IN(SELECT b.badgeID FROM tbl_attendance a, tbl_employee b " +
                        "WHERE a.EmplId = b.id AND a.date = '" + date + "' AND a.intime IS NOT NULL)) AS A ) AS A GROUP BY workarea ORDER BY FIELD(workarea, 'SMT', 'SMT-DIPPING', 'SMT-SA', 'SMT-MAINROOM')";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                string workarea = dt.Rows[i][0].ToString();
                                int total = Convert.ToInt32(dt.Rows[i][1].ToString());
                                total += total;

                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.FontName = "Times New Roman";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.FontSize = 10;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Font.Bold = true;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1).Value = "Workarea";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3).Value = ": " + workarea;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.FontName = "Times New Roman";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.FontSize = 9;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Font.Bold = true;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1).Value = "NO";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 2).Value = "Badge ID";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 3).Value = "Employee Name";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 4).Value = "Line Code";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 5).Value = "Section";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 6).Value = "Work Area";
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7).Value = "Reason";
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1), worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                worksheetAbsent.Cell(cellRowIndexStartAbsent + cellRowIndexAbsent + 1, 7).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                int cellRowIndex = cellRowIndexStartAbsent + cellRowIndexAbsent + 2;
                                int cellColumnIndex = 2;

                                Sql = "(SELECT badgeID, NAME, linecode, DESCRIPTION, workarea, reason FROM (SELECT a.badgeID, a.NAME, a.linecode, b.description, a.workarea, c.reason " +
                        "FROM tbl_employee a, tbl_masterlinecode b, tbl_attendance c WHERE a.linecode = b.name AND a.id = c.emplID AND a.status = 1 AND c.date = '" + date + "'  AND badgeID " +
                        "NOT IN(SELECT b.badgeID FROM tbl_attendance a, tbl_employee b WHERE a.EmplId = b.id AND a.date = '" + date + "' AND a.intime IS NOT NULL)) AS A " +
                        "WHERE workarea = '" + workarea + "' ) ORDER BY FIELD(DESCRIPTION, 'MGR', 'ENG', 'PC', 'PE', 'PROD', 'QC', 'STORE', 'CS'), workarea, linecode, NAME";

                                using (MySqlDataAdapter adptAbsent = new MySqlDataAdapter(Sql, myConn))
                                {
                                    DataTable dtAbsent = new DataTable();
                                    adptAbsent.Fill(dtAbsent);

                                    if (dtAbsent.Rows.Count > 0)
                                    {
                                        totalAbsent = totalAbsent + dtAbsent.Rows.Count;

                                        worksheetAbsent.Cell(3, 6).Value = "Total Absent :";
                                        worksheetAbsent.Cell(3, 7).Value = totalAbsent;
                                        worksheetAbsent.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, cellColumnIndex - 1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, cellColumnIndex - 1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

                                        // storing Each row and column value to excel sheet  
                                        for (int x = 0; x < dtAbsent.Rows.Count; x++)
                                        {
                                            for (int y = 0; y < dtAbsent.Columns.Count; y++)
                                            {
                                                worksheetAbsent.Cell(x + cellRowIndex, 1).Value = x + 1;
                                                worksheetAbsent.Cell(x + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                                if (y == 0)
                                                {
                                                    worksheetAbsent.Cell(x + cellRowIndex, y + cellColumnIndex).Value = "'" + dtAbsent.Rows[x][y].ToString();
                                                }
                                                else
                                                {
                                                    worksheetAbsent.Cell(x + cellRowIndex, y + cellColumnIndex).Value = dtAbsent.Rows[x][y].ToString();
                                                }
                                            }
                                        }
                                        int endPartAbsent = dtAbsent.Rows.Count + cellRowIndex;

                                        // setup border 
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 1), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 1), worksheetAbsent.Cell(endPartAbsent - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex, 7), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                        worksheetAbsent.Range(worksheetAbsent.Cell(endPartAbsent - 1, 1), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                                        // set value Align center
                                        worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndex - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                        cellRowIndexAbsent = endPartAbsent - 2;
                                    }
                                }
                            }
                            //worksheetAbsent.Cell(cellRowIndexAbsent + 3, 1).Value = "*Note : Break more than 90 Minutes";
                        }
                        else
                        {
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 7)).Merge();
                            worksheetAbsent.Cell(4, 1).Style.Font.FontName = "Times New Roman";
                            worksheetAbsent.Cell(4, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(4, 1).Style.Font.FontSize = 12;
                            worksheetAbsent.Cell(4, 1).Style.Font.FontColor = XLColor.Black;
                            worksheetAbsent.Cell(4, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(4, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheetAbsent.Cell(4, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheetAbsent.Cell(4, 1).Value = "No Any Data";
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }

                        //===========================================================================================================================


                        // sheet TotalLate in month     
                        var worksheetLate = workbook.Worksheets.Add("TotalLate");
                        //to hide gridlines
                        worksheetLate.ShowGridLines = false;

                        // set column width
                        worksheetLate.Columns().Width = 15;
                        worksheetLate.Column(1).Width = 5;
                        worksheetLate.Column(2).Width = 14;
                        worksheetLate.Column(3).Width = 31;
                        worksheetLate.Column(7).Width = 23;

                        worksheetLate.Rows().Height = 16.25;
                        worksheetLate.Row(1).Height = 25.5;

                        worksheetLate.PageSetup.Margins.Top = 0.5;
                        worksheetLate.PageSetup.Margins.Bottom = 0.25;
                        worksheetLate.PageSetup.Margins.Left = 0.25;
                        worksheetLate.PageSetup.Margins.Right = 0;
                        worksheetLate.PageSetup.Margins.Header = 0.5;
                        worksheetLate.PageSetup.Margins.Footer = 0.25;
                        worksheetLate.PageSetup.CenterHorizontally = true;

                        worksheetLate.Range(worksheetLate.Cell(1, 1), worksheetLate.Cell(1, 7)).Merge();
                        worksheetLate.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                        worksheetLate.Cell(1, 1).Style.Font.Bold = true;
                        worksheetLate.Cell(1, 1).Style.Font.FontSize = 20;
                        worksheetLate.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                        worksheetLate.Cell(1, 1).Style.Font.Bold = true;
                        worksheetLate.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheetLate.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        worksheetLate.Cell(1, 1).Value = "SMT TOTAL LATE PER MONTH";

                        worksheetLate.Range(worksheetLate.Cell(2, 1), worksheetLate.Cell(3, 7)).Style.Font.FontName = "Times New Roman";
                        worksheetLate.Range(worksheetLate.Cell(2, 1), worksheetLate.Cell(3, 7)).Style.Font.FontSize = 9;
                        worksheetLate.Range(worksheetLate.Cell(2, 1), worksheetLate.Cell(3, 7)).Style.Font.Bold = true;
                        worksheetLate.Range(worksheetLate.Cell(2, 4), worksheetLate.Cell(3, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        worksheetLate.Range(worksheetLate.Cell(2, 1), worksheetLate.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                        worksheetLate.Cell(2, 1).Value = "Attendance Marked At";
                        worksheetLate.Cell(2, 3).Value = ": " + date;
                        worksheetLate.Cell(2, 6).Value = "Report Date:";
                        worksheetLate.Cell(2, 7).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        int cellRowIndexStartLateMonth = 3;
                        int cellRowIndexLateMonth = 0;
                        int totalLateMonth = 0;

                        // find workarea
                        Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT b.badgeID, b.name, b.linecode, b.workarea, c.description, SUM(a.islate) AS totalLate" +
                            " FROM tbl_attendance a, tbl_employee b, tbl_masterlinecode c WHERE a.emplid = b.id AND c.name = b.linecode " +
                            "AND(a.date >= '" + startDate + "' AND a.date <= '" + endDate + "')GROUP BY b.badgeID, b.name, b.linecode, b.workarea, c.description) AS A " +
                            "GROUP BY workarea ORDER BY FIELD(workarea, 'SMT', 'SMT-DIPPING', 'SMT-SA', 'SMT-MAINROOM')";

                        using (MySqlDataAdapter adptLate = new MySqlDataAdapter(Sql, myConn))
                        {
                            DataTable dtLate = new DataTable();
                            adptLate.Fill(dtLate);

                            if (dtLate.Rows.Count > 0)
                            {
                                for (int i = 0; i < dtLate.Rows.Count; i++)
                                {
                                    string workarea = dtLate.Rows[i][0].ToString();
                                    int total = Convert.ToInt32(dtLate.Rows[i][1].ToString());
                                    total += total;

                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Font.FontName = "Times New Roman";
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Font.FontSize = 10;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Font.Bold = true;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                    worksheetLate.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1).Value = "Workarea";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3).Value = ": " + workarea;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3)).Style.Font.FontName = "Times New Roman";
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3)).Style.Font.FontSize = 9;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3)).Style.Font.Bold = true;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1).Value = "NO";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 2).Value = "Badge ID";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 3).Value = "Employee Name";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 4).Value = "Line Code";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 5).Value = "Section";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 6).Value = "Work Area";
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7).Value = "Total Late Days In a Month";
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                    worksheetLate.Range(worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1), worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                    worksheetLate.Cell(cellRowIndexStartLateMonth + cellRowIndexLateMonth + 1, 7).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                    int cellRowIndex = cellRowIndexStartLateMonth + cellRowIndexLateMonth + 2;
                                    int cellColumnIndex = 2;

                                    Sql = "(SELECT b.badgeID, b.name , b.linecode, b.workarea, c.description, SUM(a.islate) AS totalLate " +
                                        "FROM tbl_attendance a, tbl_employee b, tbl_masterlinecode c " +
                                        "WHERE a.emplid = b.id AND c.name = b.linecode AND b.workarea = '" + workarea + "' AND(a.date >= '" + startDate + "' AND a.date <= '" + endDate + "') " +
                                        "GROUP BY b.badgeID, b.name , b.linecode, b.workarea, c.description) ORDER BY totallate DESC, linecode, NAME";

                                    using (MySqlDataAdapter adpt1 = new MySqlDataAdapter(Sql, myConn))
                                    {
                                        DataTable dt1 = new DataTable();
                                        adpt1.Fill(dt1);

                                        if (dt1.Rows.Count > 0)
                                        {
                                            worksheetLate.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex, cellColumnIndex - 1), worksheetLate.Cell(dt1.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex, cellColumnIndex - 1), worksheetLate.Cell(dt1.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

                                            // storing Each row and column value to excel sheet  
                                            for (int x = 0; x < dt1.Rows.Count; x++)
                                            {
                                                for (int y = 0; y < dt1.Columns.Count; y++)
                                                {
                                                    worksheetLate.Cell(x + cellRowIndex, 1).Value = x + 1;
                                                    worksheetLate.Cell(x + cellRowIndex, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                                    if (y == 0)
                                                    {
                                                        worksheetLate.Cell(x + cellRowIndex, y + cellColumnIndex).Value = "'" + dt1.Rows[x][y].ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheetLate.Cell(x + cellRowIndex, y + cellColumnIndex).Value = dt1.Rows[x][y].ToString();
                                                    }

                                                    if (Convert.ToInt32(dt1.Rows[x][5].ToString()) >= 3)
                                                    {
                                                        worksheetLate.Cell(x + cellRowIndex, 7).Style.Fill.BackgroundColor = XLColor.Yellow;
                                                    }
                                                }
                                            }
                                            int endRow = dt1.Rows.Count + cellRowIndex;

                                            // setup border 
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex, 1), worksheetLate.Cell(endRow - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex - 1, 2), worksheetLate.Cell(endRow - 1, 7)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex, 1), worksheetLate.Cell(endRow - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex, 7), worksheetLate.Cell(endRow - 1, 7)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                            worksheetLate.Range(worksheetLate.Cell(endRow - 1, 1), worksheetLate.Cell(endRow - 1, 7)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                                            // set value Align center
                                            worksheetLate.Range(worksheetLate.Cell(cellRowIndex - 1, 2), worksheetLate.Cell(endRow - 1, 7)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                            cellRowIndexLateMonth = endRow - 2;
                                        }
                                    }
                                }
                                worksheetLate.Cell(cellRowIndexLateMonth + 3, 1).Value = "*Note : Please put high attention on employees with yellow mark ";
                            }
                            else
                            {
                                worksheetLate.Range(worksheetLate.Cell(4, 1), worksheetLate.Cell(4, 7)).Merge();
                                worksheetLate.Cell(4, 1).Style.Font.FontName = "Times New Roman";
                                worksheetLate.Cell(4, 1).Style.Font.Bold = true;
                                worksheetLate.Cell(4, 1).Style.Font.FontSize = 12;
                                worksheetLate.Cell(4, 1).Style.Font.FontColor = XLColor.Black;
                                worksheetLate.Cell(4, 1).Style.Font.Bold = true;
                                worksheetLate.Cell(4, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheetLate.Cell(4, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                worksheetLate.Cell(4, 1).Value = "No Any Data";
                                worksheetLate.Range(worksheetLate.Cell(4, 1), worksheetLate.Cell(4, 7)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                        }
                    }
                    workbook.SaveAs(directoryFile + "\\" + year + "\\Summary " + date + ".xlsx");
                }
                fileReport = directoryFile + "\\" + year + "\\Summary " + date + ".xlsx";
                System.Diagnostics.Process.Start(@"" + directoryFile + "\\" + year + "\\Summary " + date + ".xlsx");
                MessageBox.Show(this, "Excel File Success Generated", "Generate Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // tampilkan pesan error
                MessageBox.Show(ex.Message);
            }
        }


        private void timerRefresh_Tick(object sender, EventArgs e)
        {
            refresh();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            exportReport();
        }       
        // class to decide which report to export 
        private void exportReport()
        {
            string date = dateTimePickerDate.Value.ToString("yyyy-MM-dd");
            DateTime time = Convert.ToDateTime(date);
            year = time.ToString("yyyy");
            month = time.ToString("MM");
            dateReport = time.ToString("dd");

            startDate = year + "-" + month + "-1";
            endDate = year + "-" + month + "-" + dateReport;

            // get last day of month
            string lastdayOfMonth = help.TotalNumberOfDaysInMonth(Convert.ToInt32(year), Convert.ToInt32(month)).ToString();
            //check if date = end day of month
            //----save to file/xls----
            if (date == year + "-" + month + "-" + lastdayOfMonth)
            {
                ExportToExcelEndMonth();
            }
            else
            {
                ExportToExcel();
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                /// Initializes the variables to pass to the MessageBox.Show method.
                string message = "Are you sure you want to processing attendance?";
                string caption = "Process Attendance";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;
                // Displays the MessageBox.
                result = MessageBox.Show(message, caption, buttons);
                if (result == DialogResult.Yes)
                {
                    // process attendance
                    // reset data attendnace then flag data log and process transaction
                    resetAttendanceData();
                    resetFlagLog();
                    // process data
                    processTransaction();
                    MessageBox.Show(this, "Processing Attendance Complete", "Process Log Attendance", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            loadDataTransaction();
        }
    }
}
