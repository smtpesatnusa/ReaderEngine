using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;
using BIOVMyTimeSheet;
using ClosedXML.Excel;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Net;

namespace ReaderEngine
{
    public partial class frmMain : Form
    {
        private string Sql;
        MySqlConnection myConn;
        string fileReport;
        string sendto, ccto, subject, message, sendtime;

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
            catch (System.Exception ex)
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
            this.dataGridViewProcessTime.DoubleBuffered(true);
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

        private void processTransaction()
        {
            string koneksi = ConnectionDB.strProvider;
            myConn = new MySqlConnection(koneksi);
            try
            {
                myConn.Open();

                DateTime dt1 = DateTime.Today.AddDays(-2);

                string sq = "select l.*, e.id as emplid, e.workarea from tbl_log as l inner join tbl_employee as e on e.rfidno = l.rfidno " +
                   "where l.timelog>='" + dt1.ToString("yyyy-MM-dd") + "' and processed = 0 order by l.timelog, l.id desc";

                using (MySqlDataAdapter da = new MySqlDataAdapter(sq, myConn))
                {
                    var tmSheet = new Timesheets(myConn);
                    tmSheet.SetValid2Checkin(15);

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


        private bool SendSMTP(string mailaddr, string nama, string xlFile)
        {
            bool sukses = false;
            try
            {
                SmtpClient client = new SmtpClient("mail.rytechindo.com") //smtp server
                {
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential("support@rytechindo.com", "password"),
                    Port = 465,
                    EnableSsl = true
                };

                string strBody = "Dear " + nama + "\n";
                strBody += "Please find the attendance reports attached\n";
                var mailMessage = new MailMessage();
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;
                mailMessage.To.Add(mailaddr);
                mailMessage.From = new MailAddress("support@rytechindo.com", "Support", System.Text.Encoding.UTF8);
                mailMessage.Subject = "Attendance Report "+DateTime.Now.ToString() ;
                mailMessage.Body = strBody;
                mailMessage.IsBodyHtml = true;

                if (string.IsNullOrWhiteSpace(xlFile) == false)
                {
                    var attachment = new Attachment(xlFile);
                    mailMessage.Attachments.Add(attachment);
                }

                client.Send(mailMessage);
                sukses = true;
            }
            catch (Exception ex)
            {

            }
            return sukses;
        }

        private void emailDetail()
        {
            try
            {
                string koneksi = ConnectionDB.strProvider;
                myConn = new MySqlConnection(koneksi);

                string query = "SELECT sendto, cc, SUBJECT, message, sendtime FROM tbl_mastertemplateemail WHERE id = 1";
                using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, myConn))
                {
                    DataTable dset = new DataTable();
                    adpt.Fill(dset);
                    if (dset.Rows.Count > 0)
                    {
                        sendto = dset.Rows[0]["sendto"].ToString();
                        ccto = dset.Rows[0]["cc"].ToString();
                        subject = dset.Rows[0]["subject"].ToString();
                        message = dset.Rows[0]["message"].ToString();
                        sendtime = dset.Rows[0]["sendtime"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private bool SendMail(string toaddr, string ccaddr, string subjectmail, string mssgemail, string xlFile)
        {
            bool sukses = false;
            var objOutlook = new MsOutlook.Application();

            MsOutlook.MailItem newMail;
            newMail = (MsOutlook.MailItem)objOutlook.CreateItem(MsOutlook.OlItemType.olMailItem);
            MsOutlook.Accounts accounts = objOutlook.Session.Accounts;
            MsOutlook.Account acc = null;

            string accSender = "netrayaattendance@gmail.com";
            foreach (MsOutlook.Account account in accounts)
            {
                string smtpAddr = account.SmtpAddress;
                if (account.SmtpAddress.Equals(accSender, StringComparison.CurrentCultureIgnoreCase))
                {
                    acc = account;
                    break;
                }
            }

            bool bRes = false;
            if (acc != null)
            {
                //Use this account to send the e-mail. 
                newMail.SendUsingAccount = acc;
                bRes = true;
            }
            else
            {

            }

            if (bRes)
            {
                try
                {
                    {
                        newMail.To = toaddr;
                        newMail.CC = ccaddr;
                        newMail.Subject = subjectmail+" "+ DateTime.Now.AddDays(-1).ToString("MMM dd, yyyy");
                        string strBody = mssgemail;
                        newMail.Body = strBody;
                        if (string.IsNullOrWhiteSpace(xlFile) == false)
                        {
                            newMail.Attachments.Add(xlFile);
                        }
                        newMail.Send();
                        sukses = true;
                    }
                }
                catch (Exception ex)
                {
                    string retmsg = ex.Message;
                }
                finally
                {
                    newMail = null;
                    objOutlook = null;
                }
            }
            return sukses;
        }

        private void ExportToExcel()
        {
            try
            {
                string koneksi = ConnectionDB.strProvider;
                myConn = new MySqlConnection(koneksi);

                string date = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                string directoryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                directoryFile = directoryFile + "\\Attendance-SMT";
                using (var workbook = new XLWorkbook())
                {
                    // late sheet
                    var worksheet = workbook.Worksheets.Add("Late");

                    //to hide gridlines
                    worksheet.ShowGridLines = false;

                    // set column width
                    worksheet.Columns().Width = 15;
                    worksheet.Column(1).Width = 5;
                    worksheet.Column(2).Width = 14;
                    worksheet.Column(3).Width = 31;
                    worksheet.Column(9).Width = 17;

                    worksheet.Rows().Height = 16.25;
                    worksheet.Row(1).Height = 25.5;

                    // set format hour
                    worksheet.Column(7).Style.NumberFormat.Format = "hh:mm";
                    worksheet.Column(8).Style.NumberFormat.Format = "hh:mm";

                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.25;
                    worksheet.PageSetup.Margins.Left = 0.25;
                    worksheet.PageSetup.Margins.Right = 0;
                    worksheet.PageSetup.Margins.Header = 0.5;
                    worksheet.PageSetup.Margins.Footer = 0.25;
                    worksheet.PageSetup.CenterHorizontally = true;

                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 9)).Merge();
                    worksheet.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.FontName = "Times New Roman";
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.FontSize = 9;
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(2, 8), worksheet.Cell(3, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheet.Cell(2, 1).Value = "Attendance Marked At";
                    worksheet.Cell(2, 3).Value = ": " + date;
                    worksheet.Cell(2, 8).Value = "Report Date:";
                    worksheet.Cell(2, 9).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    int cellRowIndexStartLate = 3;
                    int cellRowIndexlate = 0;
                    int totalLate = 0;

                    // find workarea
                    Sql = "SELECT workarea, COUNT(*) AS total FROM (SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, " +
                        "DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, DATE_FORMAT(a.intime, '%H:%i') AS intime, " +
                        "TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus " +
                        "FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f WHERE e.id = a.emplid AND e.linecode = f.name AND " +
                        "a.date = '" + date + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' GROUP BY workarea ";

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
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Font.FontName = "Times New Roman";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 9)).Style.Font.FontSize = 9;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Font.FontSize = 10;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Font.Bold = true;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 1).Value = "Workarea :";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate, 3).Value = workarea;
                                //worksheet.Cell(3, 9).Value = "Total Late :" + total;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Value = "NO";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 2).Value = "Badge ID";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 3).Value = "Employee Name";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 4).Value = "Line Code";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 5).Value = "Section";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 6).Value = "Work Area";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 7).Value = "Schedule";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 8).Value = "Actual In";
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9).Value = "Total Late (Minute)";
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheet.Range(worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1), worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                worksheet.Cell(cellRowIndexStartLate + cellRowIndexlate + 1, 9).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                                int cellRowIndex = cellRowIndexStartLate + cellRowIndexlate + 2;
                                int cellColumnIndex = 2;
                                //worksheet.Cell(i + cellRowIndex, 1 + cellColumnIndex).Value = workarea;
                                //worksheet.Cell(i + cellRowIndex, 2 + cellColumnIndex).Value = total;

                                Sql = "SELECT badgeID, NAME, linecode, DESCRIPTION, workarea, ScheduleIn, intime, diff FROM " +
                                    "(SELECT e.badgeID, e.name, e.linecode, e.workarea, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                                    "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                                    "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f WHERE e.id = a.emplid AND e.linecode = f.name " +
                                    "AND a.date = '" + date + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' AND workarea = '" + workarea + "' ORDER BY workarea, linecode, NAME";

                                using (MySqlDataAdapter adptLateWorkarea = new MySqlDataAdapter(Sql, myConn))
                                {
                                    DataTable dtLateWorkarea = new DataTable();
                                    adptLateWorkarea.Fill(dtLateWorkarea);

                                    if (dtLateWorkarea.Rows.Count > 0)
                                    {
                                        worksheet.Cell(cellRowIndex, 3).Value = ": " + workarea;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                                        worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex - 1), worksheet.Cell(dtLateWorkarea.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

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

                                                if (Convert.ToInt32(dtLateWorkarea.Rows[x][7].ToString()) > 31)
                                                {
                                                    worksheet.Cell(x + cellRowIndex, 9).Style.Fill.BackgroundColor = XLColor.Yellow;
                                                }
                                            }
                                        }
                                        int endPart = dtLateWorkarea.Rows.Count + cellRowIndex;

                                        // setup border 
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 9)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(cellRowIndex, 9), worksheet.Cell(endPart - 1, 9)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                        worksheet.Range(worksheet.Cell(endPart - 1, 1), worksheet.Cell(endPart - 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                                        // set value Align center
                                        worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                        cellRowIndexlate = endPart - 2;
                                    }
                                }
                            }
                        }
                    }

                    // sheet Over     
                    var worksheetOver = workbook.Worksheets.Add("Overbreak");

                    //to hide gridlines
                    worksheetOver.ShowGridLines = false;

                    // set column width
                    worksheetOver.Columns().Width = 15;
                    worksheetOver.Column(1).Width = 5;
                    worksheetOver.Column(2).Width = 14;
                    worksheetOver.Column(3).Width = 31;
                    worksheetOver.Column(7).Width = 18;
                    worksheetOver.Column(8).Width = 23;

                    worksheetOver.Rows().Height = 16.25;
                    worksheetOver.Row(1).Height = 25.5;

                    worksheetOver.PageSetup.Margins.Top = 0.5;
                    worksheetOver.PageSetup.Margins.Bottom = 0.25;
                    worksheetOver.PageSetup.Margins.Left = 0.25;
                    worksheetOver.PageSetup.Margins.Right = 0;
                    worksheetOver.PageSetup.Margins.Header = 0.5;
                    worksheetOver.PageSetup.Margins.Footer = 0.25;
                    worksheetOver.PageSetup.CenterHorizontally = true;

                    worksheetOver.Range(worksheetOver.Cell(1, 1), worksheetOver.Cell(1, 8)).Merge();
                    worksheetOver.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheetOver.Cell(1, 1).Style.Font.Bold = true;
                    worksheetOver.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheetOver.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheetOver.Cell(1, 1).Style.Font.Bold = true;
                    worksheetOver.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheetOver.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheetOver.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                    worksheetOver.Range(worksheetOver.Cell(2, 1), worksheetOver.Cell(3, 8)).Style.Font.FontName = "Times New Roman";
                    worksheetOver.Range(worksheetOver.Cell(2, 1), worksheetOver.Cell(3, 8)).Style.Font.FontSize = 9;
                    worksheetOver.Range(worksheetOver.Cell(2, 1), worksheetOver.Cell(3, 8)).Style.Font.Bold = true;
                    worksheetOver.Range(worksheetOver.Cell(2, 6), worksheetOver.Cell(3, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheetOver.Range(worksheetOver.Cell(2, 1), worksheetOver.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheetOver.Cell(2, 1).Value = "Attendance Marked At :";
                    worksheetOver.Cell(2, 3).Value = date;
                    worksheetOver.Cell(2, 7).Value = "Report Date:";
                    worksheetOver.Cell(2, 8).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Font.FontName = "Times New Roman";
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Font.FontSize = 10;
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Font.Bold = true;
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheetOver.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheetOver.Cell(4, 1).Value = "NO";
                    worksheetOver.Cell(4, 2).Value = "Badge ID";
                    worksheetOver.Cell(4, 3).Value = "Employee Name";
                    worksheetOver.Cell(4, 4).Value = "Line Code";
                    worksheetOver.Cell(4, 5).Value = "Section";
                    worksheetOver.Cell(4, 6).Value = "Work Area";
                    worksheetOver.Cell(4, 7).Value = "Total Break (Minute)";
                    worksheetOver.Cell(4, 8).Value = "Total Overbreak (Minute)";
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheetOver.Range(worksheetOver.Cell(4, 1), worksheetOver.Cell(4, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    worksheetOver.Cell(4, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    worksheetOver.Cell(4, 8).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                    int cellRowIndexworksheetOver = 5;
                    int cellColumnIndexworksheetOver = 2;

                    Sql = "SELECT * FROM (SELECT b.badgeId, b.name, b.linecode, c.description, b.workarea, SUM(a.duration) AS totalBreak " +
                        ", SUM(a.duration) - 90 AS totalOver FROM tbl_durationbreak a, tbl_employee b, tbl_masterlinecode c " +
                        "WHERE a.emplid = b.id and b.linecode = c.name AND a.date = '" + date + "' " +
                        "GROUP BY b.badgeId, b.name, b.linecode, c.description,b.workarea) AS a WHERE totalbreak > 90 order by workarea, linecode, NAME";

                    using (MySqlDataAdapter adptOver = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dtOver = new DataTable();
                        adptOver.Fill(dtOver);

                        if (dtOver.Rows.Count > 0)
                        {
                            worksheetOver.Cell(3, 7).Value = "Total Overbreak :";
                            worksheetOver.Cell(3, 8).Value = dtOver.Rows.Count;
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver, cellColumnIndexworksheetOver - 1), worksheetOver.Cell(dtOver.Rows.Count + cellRowIndexworksheetOver, 9)).Style.Font.FontName = "Times New Roman";
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver, cellColumnIndexworksheetOver - 1), worksheetOver.Cell(dtOver.Rows.Count + cellRowIndexworksheetOver, 9)).Style.Font.FontSize = 9;

                            // storing Each row and column value to excel sheet  
                            for (int i = 0; i < dtOver.Rows.Count; i++)
                            {
                                for (int j = 0; j < dtOver.Columns.Count; j++)
                                {
                                    worksheetOver.Cell(i + cellRowIndexworksheetOver, 1).Value = i + 1;
                                    worksheetOver.Cell(i + cellRowIndexworksheetOver, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                    if (j == 0)
                                    {
                                        worksheetOver.Cell(i + cellRowIndexworksheetOver, j + cellColumnIndexworksheetOver).Value = "'" + dtOver.Rows[i][j].ToString();
                                    }
                                    else
                                    {
                                        worksheetOver.Cell(i + cellRowIndexworksheetOver, j + cellColumnIndexworksheetOver).Value = dtOver.Rows[i][j].ToString();
                                    }

                                    if (Convert.ToInt32(dtOver.Rows[i][5].ToString()) > 200)
                                    {
                                        worksheetOver.Range(worksheetOver.Cell(i + cellRowIndexworksheetOver, 7), worksheetOver.Cell(i + cellRowIndexworksheetOver, 8)).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    }
                                }
                            }
                            int endPartOver = dtOver.Rows.Count + cellRowIndexworksheetOver;

                            // setup border 
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver, 1), worksheetOver.Cell(endPartOver - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver - 1, 2), worksheetOver.Cell(endPartOver - 1, 8)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver, 1), worksheetOver.Cell(endPartOver - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver, 8), worksheetOver.Cell(endPartOver - 1, 8)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                            worksheetOver.Range(worksheetOver.Cell(endPartOver - 1, 1), worksheetOver.Cell(endPartOver - 1, 8)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                            // set value Align center
                            worksheetOver.Range(worksheetOver.Cell(cellRowIndexworksheetOver - 1, 2), worksheetOver.Cell(endPartOver - 1, 8)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheetOver.Cell(endPartOver + 1, 1).Value = "*Note : Break over than 90 Minutes";
                        }
                    }                

                            // sheet absent     
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

                            worksheetAbsent.Range(worksheetAbsent.Cell(1, 1), worksheetAbsent.Cell(1, 6)).Merge();
                            worksheetAbsent.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                            worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(1, 1).Style.Font.FontSize = 20;
                            worksheetAbsent.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                            worksheetAbsent.Cell(1, 1).Style.Font.Bold = true;
                            worksheetAbsent.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheetAbsent.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";

                            worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 6)).Style.Font.FontName = "Times New Roman";
                            worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 6)).Style.Font.FontSize = 9;
                            worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(3, 6)).Style.Font.Bold = true;
                            worksheetAbsent.Range(worksheetAbsent.Cell(2, 5), worksheetAbsent.Cell(3, 6)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                            worksheetAbsent.Range(worksheetAbsent.Cell(2, 1), worksheetAbsent.Cell(2, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                            worksheetAbsent.Cell(2, 1).Value = "Attendance Marked At :";
                            worksheetAbsent.Cell(2, 3).Value = date;
                            worksheetAbsent.Cell(2, 5).Value = "Report Date:";
                            worksheetAbsent.Cell(2, 6).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Font.FontName = "Times New Roman";
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Font.FontSize = 10;
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Font.Bold = true;
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheetAbsent.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            worksheetAbsent.Cell(4, 1).Value = "NO";
                            worksheetAbsent.Cell(4, 2).Value = "Badge ID";
                            worksheetAbsent.Cell(4, 3).Value = "Employee Name";
                            worksheetAbsent.Cell(4, 4).Value = "Line Code";
                            worksheetAbsent.Cell(4, 5).Value = "Section";
                            worksheetAbsent.Cell(4, 6).Value = "Work Area";
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            worksheetAbsent.Range(worksheetAbsent.Cell(4, 1), worksheetAbsent.Cell(4, 6)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                            worksheetAbsent.Cell(4, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                            worksheetAbsent.Cell(4, 6).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                            int cellRowIndexworksheetAbsent = 5;
                            int cellColumnIndexworksheetAbsent = 2;

                            Sql = "(SELECT badgeID, NAME, linecode, DESCRIPTION, workarea FROM (SELECT a.badgeID, a.NAME, a.linecode, b.description, a.workarea " +
                                "FROM tbl_employee a, tbl_masterlinecode b WHERE a.linecode = b.name AND badgeID NOT IN(SELECT b.badgeID FROM tbl_attendance a, tbl_employee b " +
                                "WHERE a.EmplId = b.id AND a.date = '" + date + "' AND a.intime IS NOT NULL)) AS A ) ORDER BY workarea, linecode, NAME";

                            using (MySqlDataAdapter adptAbsent = new MySqlDataAdapter(Sql, myConn))
                            {
                                DataTable dtAbsent = new DataTable();
                                adptAbsent.Fill(dtAbsent);

                                if (dtAbsent.Rows.Count > 0)
                                {
                                    worksheetAbsent.Cell(3, 6).Value = "Total Absent :" + dtAbsent.Rows.Count;
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, cellColumnIndexworksheetAbsent-1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndexworksheetAbsent, 9)).Style.Font.FontName = "Times New Roman";
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, cellColumnIndexworksheetAbsent-1), worksheetAbsent.Cell(dtAbsent.Rows.Count + cellRowIndexworksheetAbsent, 9)).Style.Font.FontSize = 9;

                                    // storing Each row and column value to excel sheet  
                                    for (int i = 0; i < dtAbsent.Rows.Count; i++)
                                    {
                                        for (int j = 0; j < dtAbsent.Columns.Count; j++)
                                        {
                                            worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, 1).Value = i + 1;
                                            worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                                            if (j == 0)
                                            {
                                                worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, j + cellColumnIndexworksheetAbsent).Value = "'" + dtAbsent.Rows[i][j].ToString();
                                            }
                                            else
                                            {
                                                worksheetAbsent.Cell(i + cellRowIndexworksheetAbsent, j + cellColumnIndexworksheetAbsent).Value = dtAbsent.Rows[i][j].ToString();
                                            }

                                        }
                                    }
                                    int endPartAbsent = dtAbsent.Rows.Count + cellRowIndexworksheetAbsent;

                                    // setup border 
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 1), worksheetAbsent.Cell(endPartAbsent - 1, 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 6)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 1), worksheetAbsent.Cell(endPartAbsent - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent, 6), worksheetAbsent.Cell(endPartAbsent - 1, 6)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                                    worksheetAbsent.Range(worksheetAbsent.Cell(endPartAbsent - 1, 1), worksheetAbsent.Cell(endPartAbsent - 1, 6)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                                    // set value Align center
                                    worksheetAbsent.Range(worksheetAbsent.Cell(cellRowIndexworksheetAbsent - 1, 2), worksheetAbsent.Cell(endPartAbsent - 1, 6)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                }
                            }
                            workbook.SaveAs(directoryFile + "\\" + date + "\\Summary "+date+".xlsx");
                }
                fileReport = directoryFile + "\\" + date + "\\Summary " + date + ".xlsx";
                //System.Diagnostics.Process.Start(@"" + directoryFile + "\\" + date + "\\Summary " + date + ".xlsx");
                //MessageBox.Show(this, "Excel File Success Generated", "Generate Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // tampilkan pesan error
                MessageBox.Show(ex.Message);
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
            //get email detail data send to subject etc
            emailDetail();

            string now = null;
            dateTimeNow.Text = DateTime.Now.ToString("HH:mm");
            now = DateTime.Now.ToString("HH:mm:00");

            // running process attendance
            int j = dataGridViewProcessTime.RowCount;
            if (j > 0)
            {
                for (int i = 0; i < j; i++)
                {
                    var row = dataGridViewProcessTime.Rows[i];
                    string cellString = row.Cells[0].Value.ToString();
                    if (cellString == now)
                    {
                        timer.Stop();
                        try
                        {
                            processTransaction();
                        }
                        catch (System.Exception ex)
                        {

                        }
                        finally
                        {

                        }
                        timer.Start();
                    }
                }
            }

            // send email based on setup email template
            if (DateTime.Now.ToString("HH:mm:00") == sendtime)
            {
                timer.Stop();
                //----save to file/xls----
                ExportToExcel();

                //sendto email
                if (SendMail(sendto, ccto, subject, message, fileReport))
                {
                    //MessageBox.Show("email sent!");
                }
                timer.Start();
            }
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            //if (SendSMTP("yosirwan@gmail.com", "Yos Irwan", ""))
            //{
            //    MessageBox.Show("email sent!");
            //}

            //----save to file/xls----
            ExportToExcel();
            // get detail email template
            emailDetail();
            //sendto email
            if (SendMail(sendto, ccto, subject, message, fileReport))
            {
                MessageBox.Show("email sent!");
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            timer.Stop();

            try
            {
                processTransaction();
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }
            timer.Start();
        }
    }
}
