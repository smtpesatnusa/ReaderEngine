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

                using (MySqlDataAdapter da = new MySqlDataAdapter(sq,myConn))
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

                ////----save to file/xls----
                //ExportToExcel();

                ////---send file via email---
                //SendMail("ali.sadikincom85@gmail.com", "Ali Sadikin", "e:\\Summary.xlsx");
                //SendMail("ali.sadikin@satnusa.com", "Ali Sadikin", "e:\\Summary.xlsx");

            }
            catch(System.Exception ex)
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
                    UseDefaultCredentials=false,
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
                mailMessage.Subject = "Attendance Report";
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
            catch(Exception ex){

            }
            return sukses;            
        }

        private bool SendMail(string mailaddr, string nama, string xlFile)
        {
            bool sukses = false;
            var objOutlook = new MsOutlook.Application();

            //var objNs = objOutlook.GetNamespace("MAPI");
            //objNs.Logon(null, null, true, true);

            //MsOutlook.MAPIFolder fdMail;
            //fdMail = objNs.GetDefaultFolder(MsOutlook.OlDefaultFolders.olFolderOutbox);

            MsOutlook.MailItem newMail;
            newMail = (MsOutlook.MailItem)objOutlook.CreateItem(MsOutlook.OlItemType.olMailItem);

            //newMail = (MsOutlook.MailItem)fdMail.Items.Add(MsOutlook.OlItemType.olMailItem);

            MsOutlook.Accounts accounts = objOutlook.Session.Accounts;
            MsOutlook.Account acc = null;

            string accSender = "support@rytechindo.com";
            
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

                        newMail.To = mailaddr;
                        newMail.Subject = "Attendance report";
                        string strBody = "Dear " + nama + "\n";
                        strBody += "Please find the attendance reports attached\n";
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

                string date = DateTime.Now.ToString("dd-MM-yyyy");
                string directoryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                directoryFile = directoryFile + "\\Attendance-SMT";
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Late");

                    //to hide gridlines
                    worksheet.ShowGridLines = false;

                    // set column width
                    worksheet.Columns().Width = 15;
                    worksheet.Column(1).Width = 5;
                    worksheet.Column(2).Width = 14;
                    worksheet.Column(3).Width = 31;

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

                    worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1,9)).Merge();
                    worksheet.Cell(1, 1).Style.Font.FontName = "Times New Roman";
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Font.FontSize = 20;
                    worksheet.Cell(1, 1).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(1, 1).Style.Font.Bold = true;
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Value = "SMT ATTENDANCE SUMMARY";                 

                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.FontName = "Courier New";
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.FontSize = 8;
                    worksheet.Range(worksheet.Cell(2, 1), worksheet.Cell(3, 9)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(2, 8), worksheet.Cell(3, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                    worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, 3)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheet.Cell(3, 1).Value = "Attendance Marked At :";
                    worksheet.Cell(3, 3).Value = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                    worksheet.Cell(2, 9).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Font.FontName = "Times New Roman";
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Font.FontSize = 10;
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(1, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Cell(4, 1).Value = "NO";
                    worksheet.Cell(4, 2).Value = "Badge ID";
                    worksheet.Cell(4, 3).Value = "Employee Name";
                    worksheet.Cell(4, 4).Value = "Line Code";
                    worksheet.Cell(4, 5).Value = "Section";
                    worksheet.Cell(4, 6).Value = "Schedule";
                    worksheet.Cell(4, 7).Value = "Actual In";
                    worksheet.Cell(4, 8).Value = "Diff (Minute)";
                    worksheet.Cell(4, 9).Value = "Status";
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    worksheet.Cell(4, 1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    worksheet.Cell(4, 9).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                    int cellRowIndex = 5;
                    int cellColumnIndex = 2;

                    Sql = "SELECT badgeID, NAME, linecode, DESCRIPTION, ScheduleIn, intime, diff, Sttus FROM " +
                        "(SELECT e.badgeID, e.name, e.linecode, f.description, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                        "DATE_FORMAT(a.intime, '%H:%i') AS intime, TIMESTAMPDIFF(MINUTE, a.ScheduleIn, a.intime) AS diff, " +
                        "IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus FROM tbl_attendance a, tbl_employee e, tbl_masterlinecode f WHERE e.id = a.emplid AND e.linecode = f.name " +
                        "AND a.date = '" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY linecode, NAME ";
                        
                    //"SELECT badgeID, NAME, linecode, ScheduleIn, intime, outtime, Sttus FROM (SELECT e.badgeID, e.name, e.linecode, DATE_FORMAT(a.ScheduleIn, '%H:%i') AS ScheduleIn, " +
                    //"DATE_FORMAT(a.intime, '%H:%i') AS intime, DATE_FORMAT(a.outtime, '%H:%i') AS outtime, " +
                    //"IF(a.intime > a.ScheduleIn, 'Late', 'Ontime') AS Sttus FROM tbl_attendance a, tbl_employee e WHERE e.id = a.emplid " +
                    //"AND a.date = '" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "' AND a.ScheduleIn IS NOT NULL ORDER BY a.ScheduleIn ASC) AS A WHERE Sttus = 'Late' ORDER BY intime";

                    using (MySqlDataAdapter adpt = new MySqlDataAdapter(Sql, myConn))
                    {
                        DataTable dt = new DataTable();
                        adpt.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            
                            worksheet.Cell(3, 9).Value = "Total Late :"+ dt.Rows.Count;
                            worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex-1), worksheet.Cell(dt.Rows.Count + cellRowIndex, 9)).Style.Font.FontName = "Times New Roman";
                            worksheet.Range(worksheet.Cell(cellRowIndex, cellColumnIndex-1), worksheet.Cell(dt.Rows.Count + cellRowIndex, 9)).Style.Font.FontSize = 9;

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

                                    //// set format hour
                                    //worksheet.Column(6).Style.NumberFormat.Format = "hh:mm";
                                    //worksheet.Column(7).Style.NumberFormat.Format = "hh:mm";
                                }
                            } 
                            int endPart = dt.Rows.Count + cellRowIndex;

                            // setup border 
                            worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 9)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheet.Range(worksheet.Cell(cellRowIndex, 1), worksheet.Cell(endPart - 1, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                            worksheet.Range(worksheet.Cell(cellRowIndex, 9), worksheet.Cell(endPart - 1, 9)).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                            worksheet.Range(worksheet.Cell(endPart - 1, 1), worksheet.Cell(endPart - 1, 9)).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                            // set value Align center
                            worksheet.Range(worksheet.Cell(cellRowIndex - 1, 2), worksheet.Cell(endPart - 1, 9)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            workbook.SaveAs(directoryFile + "\\" + date+"\\Summary.xlsx");
                        }
                    }
                }
                System.Diagnostics.Process.Start(@"" + directoryFile + "\\" + date+"\\Summary.xlsx");
                //MessageBox.Show(this, "Excel File Success Generated", "Generate Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // tampilkan pesan error
                //MessageBox.Show(ex.Message);
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
            string now = null;

            dateTimeNow.Text = DateTime.Now.ToString("HH:mm");
            now = DateTime.Now.ToString("HH:mm:00");

            int j = dataGridViewProcessTime.RowCount;
            if (j > 0)
            {
                for(int i = 0; i < j; i++)
                {
                    var row = dataGridViewProcessTime.Rows[i];
                    string cellString = row.Cells[0].Value.ToString();
                    if ( cellString == now)
                    {
                        timer.Stop();

                        try
                        {
                            processTransaction();
                        }
                        catch(System.Exception ex)
                        {
                            
                        }
                        finally
                        {

                        }
                        timer.Start();
                    }
                }
            }

            // running process attendance
            //try
            //{
            //    string query = "SELECT TIME FROM tbl_processTimer WHERE TIME = '" + now + "'";
            //    using (MySqlDataAdapter adpt = new MySqlDataAdapter(query, connectionDB.connection))
            //    {
            //        DataSet dset = new DataSet();
            //        adpt.Fill(dset);
            //        if (dset.Tables[0].Rows.Count > 0)
            //        {
            //            processTransaction();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    connectionDB.connection.Close();
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    connectionDB.connection.Close();
            //}
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            //if (SendMail("ali.sadikin@satnusa.com", "Ali Sadikin", ""))
            //{
            //    MessageBox.Show("email sent!");
            //}

            //if (SendSMTP("yosirwan@gmail.com", "Yos Irwan", ""))
            //{
            //    MessageBox.Show("email sent!");
            //}

            //----save to file/xls----
            ExportToExcel();

            //if (SendMail("tesemail1922@gmail.com", "Test Email", ""))
            //{
            //    MessageBox.Show("email sent!");
            //}
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
