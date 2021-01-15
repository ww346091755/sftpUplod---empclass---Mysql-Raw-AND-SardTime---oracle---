using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using MySql.Data.MySqlClient;
using System.Data.Odbc;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace sftpUplod
{
    public partial class Form1 : Form
    {
 

        public Form1()
        {
            InitializeComponent();

        }

        string strInsertError = "";

        private void Form1_Load(object sender, EventArgs e)
        {
            
            //if(File.Exists(strFilePath))
            //{
            //    txtReadPath.Text=inifile.IniReadValue("Path","ReadPath",strFilePath);
            //    txtBackupPath.Text = inifile.IniReadValue("Path", "BackupPath", strFilePath);
            //    //txtTime.Text = inifile.IniReadValue("QueryTime", "QTime", strFilePath);
            //}
            //Wip To Progress
            //Progress To dept_relation
            //Raw To Progress
            //Testemployee Empclass4 
            WriteLog("Raw、CardTime To Progress(oracle) ===================" + '\r', 1);

            

            ////InsertMySql_empClass("CSBG");
            ////InsertMySql_empClass("ASBG");
            ////InsertMySql_LmtDept("CSBG");
            ////InsertMySql_LmtDept("ASBG");
            //InsertMysql_dept_relation("CSBG");
            //InsertMysql_dept_relation("ASBG");
            //InsertMysql_dept_relation("BBBB");

            //MySql_testswipecardtimeTOprogress_CARDSR("CSBG");
            //MySql_testswipecardtimeTOprogress_CARDSR("ASBG");
            //MySql_testswipecardtimeTOprogress_CARDSR("BBBB");
            //MySql_testswipecardtimeTOprogress_CARDSR("113");
            //Mysql_rawToProgress("113");
            //Mysql_rawToProgress("ASBG");
            //Mysql_rawToProgress("BBBB");

            //InsertMySql_empClass0330("CSBG");
            //InsertMySql_empClass0330("ASBG");
            //InsertMySql_LmtDept("CSBG");
            //InsertMySql_LmtDept("ASBG");
            //InsertMySql_testemployee_SqlServer("CSBG");
            //InsertMySql_testemployee_SqlServer("ASBG");
            //InsertMySql_testemployee_Progress3("CSBG");
            //InsertMySql_testemployee_Progress3("ASBG");
     
            //InsertMysql_empClass4("CSBG");
            //InsertMysql_empClass4("ASBG");

            //InsertKK("CSBG");
            //InsertKK("ASBG");
            //haha("CSBG");
            //woriwori("ASBG");
            //Raw("ASBG");

            //nima();
            //meimei();

            #region   Raw、CardTime
            string strRaw111 = "", strRaw112 = "", strRaw113 = "", strRaw230 = "", strCardTime111 = "", strCardTime112 = "", strCardTime113 = "", strCardTime230 = "";
            bool b = true;

            try { strRaw111 = Oracle_rawToProgress(); }
            catch (System.Exception ex) { WriteLog("-->Raw111Error：" + ex.Message + '\r', 2); strRaw111 = "NG: Raw111Error：" + ex.Message; }
            //try { strRaw112 = Mysq112_rawToProgress(); }
            //catch (System.Exception ex) { WriteLog("-->Raw112Error：" + ex.Message + '\r', 2); strRaw112 = "NG: Raw112Error：" + ex.Message; }
            //try { strRaw113 = Mysq113_rawToProgress(); }
            //catch (System.Exception ex) { WriteLog("-->Raw113Error：" + ex.Message + '\r', 2); strRaw113 = "NG: Raw113Error：" + ex.Message; }
            //try { strRaw230 = Mysq230_rawToProgress(); }
            //catch (System.Exception ex) { WriteLog("-->Raw230Error：" + ex.Message + '\r', 2); strRaw230 = "NG: Raw230Error：" + ex.Message; }

            try { strCardTime111 = Oracle_testswipecardtimeTOprogress_CARDSR(); }
            catch (System.Exception ex) { WriteLog("-->CardTime111Error：" + ex.Message + '\r', 2); strCardTime111 = "NG: CardTime111Error：" + ex.Message; }
            //try { strCardTime112 = MySql112_testswipecardtimeTOprogress_CARDSR(); }
            //catch (System.Exception ex) { WriteLog("-->CardTime112Error：" + ex.Message + '\r', 2); strCardTime112 = "NG: CardTime112Error：" + ex.Message; }
            //try { strCardTime113 = MySql113_testswipecardtimeTOprogress_CARDSR(); }
            //catch (System.Exception ex) { WriteLog("-->CardTime113Error：" + ex.Message + '\r', 2); strCardTime113 = "NG: CardTime113Error：" + ex.Message; }
            //try { strCardTime230 = MySql230_testswipecardtimeTOprogress_CARDSR(); }
            //catch (System.Exception ex) { WriteLog("-->CardTime230Error：" + ex.Message + '\r', 2); strCardTime230 = "NG: CardTime230Error：" + ex.Message; }
            if (strRaw111.Substring(0, 2) == "NG") { strRaw111 = "<font color=\"red\">" + strRaw111 + "</font> "; b = false; }
            //if (strRaw112.Substring(0, 2) == "NG") { strRaw112 = "<font color=\"red\">" + strRaw112 + "</font> "; }
            //if (strRaw113.Substring(0, 2) == "NG") { strRaw113 = "<font color=\"red\">" + strRaw113 + "</font> "; }
            //if (strRaw230.Substring(0, 2) == "NG") { strRaw111 = "<font color=\"red\">" + strRaw230 + "</font> "; }
            if (strCardTime111.Substring(0, 2) == "NG") { strCardTime111 = "<font color=\"red\">" + strCardTime111 + "</font> "; b = false; }
            //if (strCardTime112.Substring(0, 2) == "NG") { strCardTime112 = "<font color=\"red\">" + strCardTime112 + "</font> "; }
            //if (strCardTime113.Substring(0, 2) == "NG") { strCardTime113 = "<font color=\"red\">" + strCardTime113 + "</font> "; }
            //if (strCardTime230.Substring(0, 2) == "NG") { strCardTime230 = "<font color=\"red\">" + strCardTime230 + "</font> "; }

            if (strInsertError != "") { b = false; };

            MailClass mail = new MailClass();

            //mail.mailFrom = "jianfeng_L@foxlink.com";
            mail.mailFrom= ConfigurationManager.AppSettings["SendAddress"]; //2018/12/27改

            //mail.mailPwd = "发送人邮箱的密码";
            mail.mailSubject = "Raw、CardTime To Progress(oracle)  " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            //mail.mailBody = strRaw111 + "<BR><BR>" + strRaw112 + "<BR><BR>" + strRaw113 + "<BR><BR>" + strRaw230 + "<BR><BR>" + strCardTime111 + "<BR><BR>" + strCardTime112 + "<BR><BR>" + strCardTime113 + "<BR><BR>" + strCardTime230 + "<BR><BR> <font color=\"red\">" + strInsertError + "</font> ";
            mail.mailBody = strRaw111 + "<BR><BR>" + strCardTime111 + "<BR><BR> <font color=\"red\">" + strInsertError + "</font> ";
            mail.isbodyHtml = true;    //是否是HTML

            //mail.host = "192.168.64.249";//SMTP服务器IP地址
            mail.host= ConfigurationManager.AppSettings["HOST"];//2018/12/27改

            //mail.mailToArray = new string[] { "jianfeng_L@foxlink.com" };//接收者邮件集合
            string[] strReveiver = ConfigurationManager.AppSettings["Reveiver"].Split(';');//2018/12/27改
            mail.mailToArray = strReveiver;//2018/12/27改


            mail.important = true;//邮件优先级默认为Normal
            if (!b)
            {
                mail.important = false;//如果有错误，邮件优先级为High
            }
            if (mail.Send())
            {
                WriteLog("邮件发送成功" + '\r', 1);//发送成功则提示返回当前页面..;

            }
            else
            {
                WriteLog("发送邮件出错,原因：" + '\r', 2);
            }
            this.Close();
            #endregion

        }

        /// <summary>
        /// 开始按钮逻辑..
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            

            if (btnStart.Text.Equals("Start"))
            {
                //检查路径是否正确
                //if (!CheckPath()) return;

                btnStart.Text = "Pause";
                WriteLog("01:00 InsertMySQL，10:00 InsertProgress" + '\r',1);
                timer1.Enabled = true;
            }
            else
            {
                btnStart.Text = "Start";
                WriteLog("Pause Read..." + '\r',1);
                timer1.Enabled = false;
            }
        }




        /// <summary>
        /// 輸出訊息在畫面上,並寫入Log檔....
        /// </summary>
        /// <param name="intColorType">0:Gray,1:Blue,2:Red</param>
        private void WriteLog(String sMessage, Int32 intColorType = 0)
        {
            DateTime currenttime = System.DateTime.Now;
            string sDateTime = string.Format("{0:yyyy-MM-dd HH:mm:ss}", currenttime);
            RichTxtLog.AppendText(sDateTime + " => " + sMessage);
            //RichTxtLog.Text=sMessage+'\r'+RichTxtLog.Text;
            RichTxtLog.Select(RichTxtLog.TextLength - sMessage.Length, RichTxtLog.TextLength);
            switch (intColorType)
            {
                //case 0:
                //    //一般訊息
                //    RichTxtLog.SelectionColor = Color.Gray;
                //    break;
                case 1:
                    //上傳成功
                    RichTxtLog.SelectionColor = Color.Green;
                    break;
                case 2:
                    //出錯
                    RichTxtLog.SelectionColor = Color.Red;
                    break;
            }
            ////保留最新的100筆訊息
            //String[] lines = new String[100];
            //if (RichTxtLog.Lines.Length > lines.Length)
            //{
            //    Array.Copy(RichTxtLog.Lines, lines, lines.Length);
            //    this.RichTxtLog.Lines = lines;
            //}

            if (RichTxtLog.Lines.Length > 100)
            {
                string[] sLines = RichTxtLog.Lines;
                string[] sNewLines = new string[sLines.Length - 100];

                Array.Copy(sLines, 100, sNewLines, 0, sNewLines.Length);
                RichTxtLog.Lines = sNewLines;
            }

            //create log folder
            string sLogPath = System.Environment.CurrentDirectory + "\\sftplog";
            if (!Directory.Exists(sLogPath))
            {
                Directory.CreateDirectory(sLogPath);
            }
            //write log to local
            StreamWriter sw = null;
            try
            {
                string sTitle = string.Format("{0:yyyyMMdd}", currenttime);
                string LogStr = currenttime.ToString("HH:mm:ss:fffffff") + " " + sMessage;
                sw = new StreamWriter(sLogPath + "\\" + sTitle + ".txt", true);
                sw.WriteLine(LogStr);
                //sw.WriteLine('\r');
            }
            catch
            {
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
            }

        }

        private void RichTxtLog_TextChanged(object sender, EventArgs e)
        {

            RichTxtLog.SelectionStart = RichTxtLog.Text.Length; //Set the current caret position at the end     
            RichTxtLog.ScrollToCaret();
            //RichTxtLog.ScrollToCaret(); //Now scroll it automatically
        }

        


        /// <summary>
        /// 定时器.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            //timer1.Enabled = false;
            

            //#region 01:00 
            //if (DateTime.Now.ToString("HH:mm") == "01:00")
            //{

            //    InsertMySql_empClass("CSBG");
            //    InsertMySql_empClass("ASBG");
            //    InsertMySql_LmtDept("CSBG");
            //    InsertMySql_LmtDept("ASBG");





            //}
            //#endregion

            //if (DateTime.Now.ToString("HH:mm") == "10:00")
            //{
            //    MySql_testswipecardtimeTOprogress_CARDSR("CSBG");
            //    MySql_testswipecardtimeTOprogress_CARDSR("ASBG");
            //}



           
            //timer1.Enabled = true;


        }

        private void MySql_testswipecardtimeTOprogress_CARDSR(string strBG)
        {
            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
  FROM  swipecard.testswipecardtime cardtime
 WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
                   DATE_SUB(CURDATE(), INTERVAL 3 DAY))
            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
                       DATE_SUB(CURDATE(), INTERVAL 3 DAY)
                AND cardtime.SwipeCardTime IS NULL)
            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = DATE_SUB(CURDATE(), INTERVAL 2 DAY)
                AND shift = 'N'))";

//            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
//  FROM  swipecard.testswipecardtime cardtime
// WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
//                   DATE_SUB(CURDATE(), INTERVAL 1 DAY))
//            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
//                       DATE_SUB(CURDATE(), INTERVAL 1 DAY)
//                AND cardtime.SwipeCardTime IS NULL)
//            OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
//                AND shift = 'N'))";



            string sql2 = @"SELECT emp.ID,
       cardtime.RecordID,
       cardtime.CardID,
       cardtime.name,
       cardtime.SwipeCardTime,
       cardtime.SwipeCardTime2,
       cardtime.CheckState,
       cardtime.PROD_LINE_CODE,
       cardtime.WorkshopNo,
       cardtime.PRIMARY_ITEM_NO,
       cardtime.RC_NO,
       cardtime.Shift,
       cardtime.OvertimeState,
       cardtime.overtimeType,
       cardtime.overtimeCal
  FROM swipecard.testemployee emp, swipecard.testswipecardtime cardtime
 WHERE     emp.cardid = cardtime.CardID
       AND (   (DATE_FORMAT(swipecardtime, '%Y%m%d') =
                   DATE_SUB(CURDATE(), INTERVAL 3 DAY))
            OR (    DATE_FORMAT(swipecardtime2, '%Y%m%d') =
                       DATE_SUB(CURDATE(), INTERVAL 3 DAY)
                AND cardtime.SwipeCardTime IS NULL)
            OR (    DATE_FORMAT(swipecardtime2, '%Y%m%d') = DATE_SUB(CURDATE(), INTERVAL 2 DAY)
                AND shift = 'N'))";
            //string sql2 = @"";
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDBC(sql);
                if (DT2 == null) { WriteLog( "-->Query 111wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog( "-->Query 111wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog( "-->Query 111wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 
                
                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                            ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                            sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                            sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                            string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                                 + dr[1].ToString() + "','"
                                                                                 + dr[2].ToString() + "','"
                                                                                 + dr[3].ToString() + "',"
                                                                                 + sa + ","
                                                                                 + sb + ",'"
                                                                                 + dr[6].ToString() + "','"
                                                                                 + dr[7].ToString() + "','"
                                                                                 + dr[8].ToString() + "','"
                                                                                 + dr[9].ToString() + "','"
                                                                                 + dr[10].ToString() + "','"
                                                                                 + dr[11].ToString() + "','"
                                                                                 + dr[12].ToString() + "','"
                                                                                 + dr[13].ToString() + "','"
                                                                                 + dr[14].ToString() + "')";

                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();
                        //}

                    }


                    tx.Commit();
                    
                    WriteLog( "-->111WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->111WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
            }
            else if (strBG == "ASBG")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
                if (DT2 == null) { WriteLog( "-->Query 112wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog( "-->Query 112wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog( "-->Query 112wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键
                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                            ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                            sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                            sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                            string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                                 + dr[1].ToString() + "','"
                                                                                 + dr[2].ToString() + "','"
                                                                                 + dr[3].ToString() + "',"
                                                                                 + sa + ","
                                                                                 + sb + ",'"
                                                                                 + dr[6].ToString() + "','"
                                                                                 + dr[7].ToString() + "','"
                                                                                 + dr[8].ToString() + "','"
                                                                                 + dr[9].ToString() + "','"
                                                                                 + dr[10].ToString() + "','"
                                                                                 + dr[11].ToString() + "','"
                                                                                 + dr[12].ToString() + "','"
                                                                                 + dr[13].ToString() + "','"
                                                                                 + dr[14].ToString() + "')";

                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();
                        //}
                    }


                    tx.Commit();
                    WriteLog( "-->112WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->112WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
            }
            else if (strBG == "BBBB")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDB65_230(sql2);
                if (DT2 == null) { WriteLog("-->Query 230wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 230wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog("-->Query 230wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 
                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    //string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                         //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                         //if (Progressdr == null)
                         //{
                             sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                             sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                             string strSQL = "insert into pub.CARDSR  VALUES  ('" + dr[0].ToString() + "','"
                                                                                  + dr[1].ToString() + "','"
                                                                                  + dr[2].ToString() + "','"
                                                                                  + dr[3].ToString() + "',"
                                                                                  + sa + ","
                                                                                  + sb + ",'"
                                                                                  + dr[6].ToString() + "','"
                                                                                  + dr[7].ToString() + "','"
                                                                                  + dr[8].ToString() + "','"
                                                                                  + dr[9].ToString() + "','"
                                                                                  + dr[10].ToString() + "','"
                                                                                  + dr[11].ToString() + "','"
                                                                                  + dr[12].ToString() + "','"
                                                                                  + dr[13].ToString() + "','"
                                                                                  + dr[14].ToString() + "')";

                             cmd.CommandText = strSQL;
                             cmd.ExecuteNonQuery();
                        // }

                    }


                    tx.Commit();
                    WriteLog("-->230WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->230WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                
 
            }
            if (strBG == "113")
            {
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
                DataTable DT2 = MySqlHelp.QueryDb113(sql);
                if (DT2 == null) { WriteLog("-->Query 113wipecardError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 113wipecardSum:0" + '\r', 2); return; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
                WriteLog("-->Query 113wipecardSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                        ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                        sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                             + dr[1].ToString() + "','"
                                                                             + dr[2].ToString() + "','"
                                                                             + dr[3].ToString() + "',"
                                                                             + sa + ","
                                                                             + sb + ",'"
                                                                             + dr[6].ToString() + "','"
                                                                             + dr[7].ToString() + "','"
                                                                             + dr[8].ToString() + "','"
                                                                             + dr[9].ToString() + "','"
                                                                             + dr[10].ToString() + "','"
                                                                             + dr[11].ToString() + "','"
                                                                             + dr[12].ToString() + "','"
                                                                             + dr[13].ToString() + "','"
                                                                             + dr[14].ToString() + "')";

                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //}

                    }


                    tx.Commit();

                    WriteLog("-->113WipToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->113WipToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
            }
 
        }

        private void Mysql_rawToProgress(string strBG)
        {
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       update_time
  FROM swipecard.raw_record
 WHERE     record_status!='4'";
//            string sql = @"SELECT id,
//       cardid,
//       swipecardtime,
//       update_time
//  FROM swipecard.raw_record
// WHERE     swipecardtime >= DATE_ADD(DATE_SUB(CURDATE(), INTERVAL 1 DAY), INTERVAL 10 HOUR)
//       AND swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR) AND record_status!='4'";

//            string sql = @"SELECT id,
//       cardid,
//       swipecardtime,
//       update_time
//  FROM swipecard.raw_record
// WHERE swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR)";
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                #region CSBG
                DataTable DT2 = MySqlHelp.QueryDBC(sql);
                if (DT2 == null) { WriteLog("-->Query 111RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 111RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 111RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                       + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->111RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    WriteLog("-->111RawToProgressError" + ex.Message + '\r', 2);
                    tx.Rollback();
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
                if (DT2 == null) { WriteLog("-->Query 112RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 112RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 112RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                       + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->112RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->112RawToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }
            else if (strBG == "BBBB")
            {
                #region BBBB
                DataTable DT2 = MySqlHelp.QueryDB65_230(sql);
                if (DT2 == null) { WriteLog("-->Query 230RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 230RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 230RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                       + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->230RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->230RawToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }
            else if (strBG == "113")
            {
                #region 113
                DataTable DT2 = MySqlHelp.QueryDb113(sql);
                if (DT2 == null) { WriteLog("-->Query 113RawError" + '\r', 2); return; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query 113RawSum:0" + '\r', 2); return; }
                WriteLog("-->Query 113RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                       + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tx.Commit();

                    WriteLog("-->113RawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->113RawToProgressError" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion
            }
 
        }

        private string Oracle_rawToProgress()
        {
            string status=  ConfigurationManager.AppSettings["Status"];//2019/05/16改
                                                                       //           string sql = @"SELECT id,
                                                                       //      cardid,
                                                                       //      swipecardtime,
                                                                       //      sysdate update_time 
                                                                       // FROM swipe.raw_record
                                                                       //WHERE UPDATE_TIME >=
                                                                       //         to_date(to_char(TRUNC(sysdate-1),'yyyy/mm/dd')|| ' 09:00:00','yyyy/mm/dd HH24:MI:SS')
                                                                       //      AND UPDATE_TIME <
                                                                       //             to_date(to_char(TRUNC(sysdate),'yyyy/mm/dd')|| ' 09:00:00','yyyy/mm/dd HH24:MI:SS') AND record_status not in (" + status + ")";
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       sysdate update_time 
  FROM swipe.raw_record
 WHERE UPDATE_TIME >=
          to_date(to_char(TRUNC(sysdate-1),'yyyy/mm/dd')|| ' 09:00:00','yyyy/mm/dd HH24:MI:SS')
       AND UPDATE_TIME <
              to_date(to_char(TRUNC(sysdate),'yyyy/mm/dd')|| ' 09:00:00','yyyy/mm/dd HH24:MI:SS') AND record_status not in (" + status + ")";
            //          string sql = @"              SELECT id,
            //     cardid,
            //     swipecardtime,
            //     sysdate update_time
            //FROM swipe.raw_record where RECORD_STATUS='10'";
            string result = "NG";//enen

            OracleHelp OraclelHelp = new OracleHelp();
            ProgressHelp ProHelp = new ProgressHelp();
           
                #region CSBG
                DataTable DT2 = OraclelHelp.QueryDbCSBG(sql);
                if (DT2 == null) { WriteLog("-->Query OracleRawError" + '\r', 2); return result = "NG: Query OracleRawError"; }
                else if (DT2.Rows.Count == 0) { WriteLog("-->Query OracleRawSum:0" + '\r', 2); return result = "NG: Query OracleRawSum:0"; }
                WriteLog("-->Query OracleRawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=DKHR;UID=csruser;PWD=csruser01";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {
                    string sa = "";
                    string sb = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                    //DateTime dtSWIPECARDTIME = Convert.ToDateTime(dr["swipecardtime"].ToString());//刷卡时间
                    //DateTime dtToadyTen = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 10:00:00");//设定一个时间，当天的10:00:00              
                    //int compNum = DateTime.Compare(dtSWIPECARDTIME, dtToadyTen);                    
                    //if (compNum > 0)//比较刷卡时间会不会大于当天的10点。
                    //{
                    //    continue;//刷卡时间大于当天的10点钟，不插入progress，继续下一循环。
                    //}


                    sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                        sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                        string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                       + sb + ",'"
                                                                                                                       + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                       + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (System.Exception ex)
                        {
                            string s = ex.Message.Trim().Substring(ex.Message.Trim().Length - 6, 5);
                            if (s != "16949")
                            {
                                WriteLog("-->OracleRawInsertError " + ex.Message + " " + strSQL + '\r', 2);
                                strInsertError += "OracleRawInsertError" + ex.Message + '\r';
                            }
                            continue;
                        }
                    }
                    tx.Commit();
                    result = "OK: OracleRawToProgress";
                    WriteLog("-->OracleRawToProgressOK" + '\r', 1);
                }
                catch (System.Exception ex)
                {
                    WriteLog("-->OracleRawToProgressErrorRollback" + ex.Message + '\r', 2);
                    result = "NG: OracleRawToProgressErrorRollback" + ex.Message;
                    tx.Rollback();
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                #endregion

                return result;
            
            

        }
        private string Mysq112_rawToProgress()
        {
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       update_time
  FROM swipecard.raw_record
 WHERE     swipecardtime >= DATE_ADD(DATE_SUB(CURDATE(), INTERVAL 1 DAY), INTERVAL 10 HOUR)
       AND swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR) AND record_status!='4'";
            string result = "NG";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            #region CSBG
            DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
            if (DT2 == null) { WriteLog("-->Query 112RawError" + '\r', 2); return result = "NG: Query 112RawError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 112RawSum:0" + '\r', 2); return result = "NG: Query 112RawSum:0"; }
            WriteLog("-->Query 112RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {
                string sa = "";
                string sb = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                    sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                    string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                   + sb + ",'"
                                                                                                                   + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                   + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        string s = ex.Message.Trim().Substring(ex.Message.Trim().Length - 6, 5);
                        if (s != "16949")
                        {
                            WriteLog("-->111RawInsertError " + ex.Message + " " + strSQL + '\r', 2);
                            strInsertError += "111RawInsertError" + ex.Message + '\r';
                        }
                        continue;
                    }
                }
                tx.Commit();
                result = "OK: 112RawToProgress";
                WriteLog("-->112RawToProgressOK" + '\r', 1);
            }
            catch (System.Exception ex)
            {
                WriteLog("-->112RawToProgressError" + ex.Message + '\r', 2);
                result = "NG: 112RawToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            #endregion

            return result;



        }
        private string Mysq113_rawToProgress()
        {
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       update_time
  FROM swipecard.raw_record
 WHERE     swipecardtime >= DATE_ADD(DATE_SUB(CURDATE(), INTERVAL 1 DAY), INTERVAL 10 HOUR)
       AND swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR) AND record_status!='4'";
            string result = "NG";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            #region CSBG
            DataTable DT2 = MySqlHelp.QueryDb113(sql);
            if (DT2 == null) { WriteLog("-->Query 113RawError" + '\r', 2); return result = "NG: Query 113RawError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 113RawSum:0" + '\r', 2); return result = "NG: Query 113RawSum:0"; }
            WriteLog("-->Query 113RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {
                string sa = "";
                string sb = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                    sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                    string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                   + sb + ",'"
                                                                                                                   + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                   + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        string s = ex.Message.Trim().Substring(ex.Message.Trim().Length - 6, 5);
                        if (s != "16949")
                        {
                            WriteLog("-->111RawInsertError " + ex.Message + " " + strSQL + '\r', 2);
                            strInsertError += "111RawInsertError" + ex.Message + '\r';
                        }
                        continue;
                    }
                }
                tx.Commit();
                result = "OK: 113RawToProgress";
                WriteLog("-->113RawToProgressOK" + '\r', 1);
            }
            catch (System.Exception ex)
            {
                WriteLog("-->113RawToProgressError" + ex.Message + '\r', 2);
                result = "NG: 113RawToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            #endregion

            return result;



        }
        private string Mysq230_rawToProgress()
        {
            string sql = @"SELECT id,
       cardid,
       swipecardtime,
       update_time
  FROM swipecard.raw_record
 WHERE     swipecardtime >= DATE_ADD(DATE_SUB(CURDATE(), INTERVAL 1 DAY), INTERVAL 10 HOUR)
       AND swipecardtime < DATE_ADD(CURDATE(), INTERVAL 10 HOUR) AND record_status!='4'";
            string result = "NG";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            #region CSBG
            DataTable DT2 = MySqlHelp.QueryDB65_230(sql);
            if (DT2 == null) { WriteLog("-->Query 230RawError" + '\r', 2); return result = "NG: Query 230RawError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 230RawSum:0" + '\r', 2); return result = "NG: Query 230RawSum:0"; }
            WriteLog("-->Query 230RawSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {
                string sa = "";
                string sb = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    sa = dr["id"].Equals(DBNull.Value) ? "NULL" : "'" + dr["id"].ToString() + "'";
                    sb = dr["cardid"].Equals(DBNull.Value) ? "NULL" : "'" + dr["cardid"].ToString() + "'";
                    string strSQL = "insert into pub.CARDSR2 (EMP_CD,CardID,SwipeCardTime,UPDATE_TIME)  VALUES  (" + sa + ","
                                                                                                                   + sb + ",'"
                                                                                                                   + Convert.ToDateTime(dr["swipecardtime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "','"
                                                                                                                   + Convert.ToDateTime(dr["update_time"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        string s = ex.Message.Trim().Substring(ex.Message.Trim().Length - 6, 5);
                        if (s != "16949")
                        {
                            WriteLog("-->111RawInsertError " + ex.Message + " " + strSQL + '\r', 2);
                            strInsertError += "111RawInsertError" + ex.Message + '\r';
                        }
                        continue;
                    }
                }
                tx.Commit();
                result = "OK: 230RawToProgress";
                WriteLog("-->230RawToProgressOK" + '\r', 1);
            }
            catch (System.Exception ex)
            {
                WriteLog("-->230RawToProgressError" + ex.Message + '\r', 2);
                result = "NG: 230RawToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            #endregion

            return result;



        }
        private string Oracle_testswipecardtimeTOprogress_CARDSR()
        {
            //            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
            //              FROM  swipecard.testswipecardtime cardtime
            //             WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
            //                               DATE_SUB(CURDATE(), INTERVAL 1 DAY))
            //                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
            //                                   DATE_SUB(CURDATE(), INTERVAL 1 DAY)
            //                            AND cardtime.SwipeCardTime IS NULL)
            //                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
            //                            AND shift = 'N'))";
            //          string sql = @"SELECT  T.EMP_ID,
            //       T.SWIPECARDTIME,
            //       T.SWIPECARDTIME2,
            //       T.CHECKSTATE,
            //       T.PROD_LINE_CODE,
            //       T.WORKSHOPNO,
            //       T.PRIMARY_ITEM_NO,
            //       T.RC_NO,
            //       T.SHIFT
            //FROM SWIPE.CSR_SWIPECARDTIME T WHERE T.SWIPE_DATE=TO_CHAR(SYSDATE-2,'YYYY-MM-DD')";
            string sql = @"SELECT  T.EMP_ID,
         T.SWIPECARDTIME,
         T.SWIPECARDTIME2,
         T.CHECKSTATE,
         T.PROD_LINE_CODE,
         T.WORKSHOPNO,
         T.PRIMARY_ITEM_NO,
         T.RC_NO,
         T.SHIFT
  FROM SWIPE.CSR_SWIPECARDTIME T WHERE T.SWIPE_DATE=TO_CHAR(SYSDATE-2,'YYYY-MM-DD')";

            OracleHelp OraclelHelp = new OracleHelp();
            ProgressHelp ProHelp = new ProgressHelp();
            string result = "NG";
            #region CSBG
           
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
                //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
            DataTable DT2 = OraclelHelp.QueryDbCSBG(sql);
            if (DT2 == null) { WriteLog("-->Query OracleCardTimeError" + '\r', 2); return result = "NG: Query OracleCardTimeError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query OracleCardTimeSum:0" + '\r', 2); return result = "NG: Query OracleCardTimeSum:0"; }

                //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 .
            WriteLog("-->Query OracleCardTimeSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //DataTable DT3 = ProHelp.QueryProgress(sql2);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
                //    return;
                //}
                //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

                //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                string ControlString1 = "DSN=DKHR;UID=csruser;PWD=csruser01";
                OdbcConnection odbcCon = new OdbcConnection(ControlString1);
                odbcCon.Open();
                OdbcTransaction tx = odbcCon.BeginTransaction();
                OdbcCommand cmd = new OdbcCommand();
                cmd.Connection = odbcCon;
                cmd.Transaction = tx;
                try
                {

                    string sa = "";
                    string sb = "";
                    string ss = "";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                        //if (Progressdr == null)
                        //{
                        //ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[2].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                        string strSQL = "insert into pub.CARDSR (EMP_CD,swipecardtime,swipecardtime2,checkstate,prod_line_code,workshopno,primary_item_no,rc_no,shift)  VALUES  ('" 
                                                                             
                                                                             + dr[0].ToString() + "',"
                                                                         
                                                                             + sa + ","
                                                                             + sb + ",'"
                                                                             + dr[3].ToString() + "','"
                                                                             + dr[4].ToString() + "','"
                                                                             + dr[5].ToString() + "','"
                                                                             + dr[6].ToString() + "','"
                                                                             + dr[7].ToString() + "','"
                                                                            
                                                                             + dr[8].ToString() + "')";

                        cmd.CommandText = strSQL;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (System.Exception ex)
                        {
                            WriteLog("-->OracleCardTimeInsertError " + strSQL + '\r', 2);
                            strInsertError += "OracleCardTimeInsertError " + ex.Message + '\r';
                            continue;
                        }
                        //}

                    }


                    tx.Commit();
                    result = "OK: OracleCardTimeToProgress";
                    WriteLog("-->OracleCardTimeToProgressOK" + '\r', 1);



                }
                catch (System.Exception ex)
                {
                    WriteLog("-->OracleCardTimeToProgressErrorRollback" + ex.Message + '\r', 2);
                    result = "NG: OracleCardTimeToProgressErrorRollback" + ex.Message;
                    tx.Rollback();  
                }
                finally
                {
                    tx.Dispose();
                    odbcCon.Close();

                }
                return result;
            #endregion
            

        }
        private string MySql112_testswipecardtimeTOprogress_CARDSR()
        {
            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
              FROM  swipecard.testswipecardtime cardtime
             WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
                               DATE_SUB(CURDATE(), INTERVAL 1 DAY))
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
                                   DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                            AND cardtime.SwipeCardTime IS NULL)
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
                            AND shift = 'N'))";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string result = "NG";
            #region 112

            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
            DataTable DT2 = MySqlHelp.QueryDbAsbg(sql);
            if (DT2 == null) { WriteLog("-->Query 112wipecardError" + '\r', 2); return result = "NG: Query 112CardTimeError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 112wipecardSum:0" + '\r', 2); return result = "NG: Query 112CardTimeSum:0"; }

            //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
            WriteLog("-->Query 112CardTimeSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            //DataTable DT3 = ProHelp.QueryProgress(sql2);
            //if (DT3 == null)
            //{
            //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
            //    return;
            //}
            //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

            //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {

                string sa = "";
                string sb = "";
                string ss = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                    //if (Progressdr == null)
                    //{
                    ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                    sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                         + dr[1].ToString() + "','"
                                                                         + dr[2].ToString() + "','"
                                                                         + dr[3].ToString() + "',"
                                                                         + sa + ","
                                                                         + sb + ",'"
                                                                         + dr[6].ToString() + "','"
                                                                         + dr[7].ToString() + "','"
                                                                         + dr[8].ToString() + "','"
                                                                         + dr[9].ToString() + "','"
                                                                         + dr[10].ToString() + "','"
                                                                         + dr[11].ToString() + "','"
                                                                         + dr[12].ToString() + "','"
                                                                         + dr[13].ToString() + "','"
                                                                         + dr[14].ToString() + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        WriteLog("-->112CardTimeInsertError " + strSQL + '\r', 2);
                        strInsertError += "112CardTimeInsertError "+ex.Message + '\r';
                        continue;
                    }
                    //}

                }


                tx.Commit();
                result = "OK: 112CardTimeToProgress";
                WriteLog("-->112CardTimeToProgressOK" + '\r', 1);



            }
            catch (System.Exception ex)
            {
                WriteLog("-->112CardTimeToProgressError" + ex.Message + '\r', 2);
                result = "NG: 112CardTimeToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            return result;
            #endregion


        }
        private string MySql113_testswipecardtimeTOprogress_CARDSR()
        {
            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
              FROM  swipecard.testswipecardtime cardtime
             WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
                               DATE_SUB(CURDATE(), INTERVAL 1 DAY))
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
                                   DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                            AND cardtime.SwipeCardTime IS NULL)
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
                            AND shift = 'N'))";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string result = "NG";
            #region 113

            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
            DataTable DT2 = MySqlHelp.QueryDb113(sql);
            if (DT2 == null) { WriteLog("-->Query 113wipecardError" + '\r', 2); return result = "NG: Query 113CardTimeError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 113wipecardSum:0" + '\r', 2); return result = "NG: Query 113CardTimeSum:0"; }

            //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
            WriteLog("-->Query 113CardTimeSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            //DataTable DT3 = ProHelp.QueryProgress(sql2);
            //if (DT3 == null)
            //{
            //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
            //    return;
            //}
            //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

            //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {

                string sa = "";
                string sb = "";
                string ss = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                    //if (Progressdr == null)
                    //{
                    ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                    sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                         + dr[1].ToString() + "','"
                                                                         + dr[2].ToString() + "','"
                                                                         + dr[3].ToString() + "',"
                                                                         + sa + ","
                                                                         + sb + ",'"
                                                                         + dr[6].ToString() + "','"
                                                                         + dr[7].ToString() + "','"
                                                                         + dr[8].ToString() + "','"
                                                                         + dr[9].ToString() + "','"
                                                                         + dr[10].ToString() + "','"
                                                                         + dr[11].ToString() + "','"
                                                                         + dr[12].ToString() + "','"
                                                                         + dr[13].ToString() + "','"
                                                                         + dr[14].ToString() + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        WriteLog("-->113CardTimeInsertError " + strSQL + '\r', 2);
                        strInsertError += "113CardTimeInsertError "+ex.Message + '\r';
                        continue;
                    }
                    //}

                }


                tx.Commit();
                result = "OK: 113CardTimeToProgress";
                WriteLog("-->113CardTimeToProgressOK" + '\r', 1);



            }
            catch (System.Exception ex)
            {
                WriteLog("-->113CardTimeToProgressError" + ex.Message + '\r', 2);
                result = "NG: 113CardTimeToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            return result;
            #endregion


        }
        private string MySql230_testswipecardtimeTOprogress_CARDSR()
        {
            string sql = @"SELECT cardtime.ID , cardtime.RecordID, cardtime.CardID, cardtime.name, cardtime.SwipeCardTime, cardtime.SwipeCardTime2, cardtime.CheckState, cardtime.PROD_LINE_CODE, cardtime.WorkshopNo, cardtime.PRIMARY_ITEM_NO, cardtime.RC_NO, cardtime.Shift, cardtime.OvertimeState, cardtime.overtimeType, cardtime.overtimeCal 
              FROM  swipecard.testswipecardtime cardtime
             WHERE   (   (DATE_FORMAT(cardtime.swipecardtime, '%Y%m%d') =
                               DATE_SUB(CURDATE(), INTERVAL 1 DAY))
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') =
                                   DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                            AND cardtime.SwipeCardTime IS NULL)
                        OR (    DATE_FORMAT(cardtime.swipecardtime2, '%Y%m%d') = CURDATE()
                            AND shift = 'N'))";

            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string result = "NG";
            #region 230

            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and cardtime.SwipeCardTime>='2017/8/15 00:00' and  cardtime.SwipeCardTime<'2017/8/16 00:00'";
            //string sql = "select emp.ID,cardtime.* from swipecard.testemployee emp, swipecard.testswipecardtime cardtime where emp.cardid = cardtime.CardID and ((DATE_FORMAT(swipecardtime, '%Y%m%d')=curdate()-1) or (DATE_FORMAT(swipecardtime2, '%Y%m%d')=curdate()-1 and cardtime.Shift='N' and cardtime.SwipeCardTime is null))";
            DataTable DT2 = MySqlHelp.QueryDB65_230(sql);
            if (DT2 == null) { WriteLog("-->Query 230wipecardError" + '\r', 2); return result = "NG: Query 230CardTimeError"; }
            else if (DT2.Rows.Count == 0) { WriteLog("-->Query 230wipecardSum:0" + '\r', 2); return result = "NG: Query 230CardTimeSum:0"; }

            //DT2.PrimaryKey = new DataColumn[] { DT2.Columns["RecordID"] };//设置RecordID为主键 
            WriteLog("-->Query 230CardTimeSum:" + DT2.Rows.Count.ToString() + '\r', 1);

            //DataTable DT3 = ProHelp.QueryProgress(sql2);
            //if (DT3 == null)
            //{
            //    WriteLog("-->Query ProgressCardsrError" + '\r', 2);
            //    return;
            //}
            //DT3.PrimaryKey = new DataColumn[] { DT3.Columns["RecordID"] };//设置RecordID为主键 

            //WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

            string ControlString1 = "";
            OdbcConnection odbcCon = new OdbcConnection(ControlString1);
            odbcCon.Open();
            OdbcTransaction tx = odbcCon.BeginTransaction();
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = odbcCon;
            cmd.Transaction = tx;
            try
            {

                string sa = "";
                string sb = "";
                string ss = "";
                foreach (DataRow dr in DT2.Rows)
                {
                    //DataRow Progressdr = DT3.Rows.Find(dr["RecordID"].ToString());
                    //if (Progressdr == null)
                    //{
                    ss = dr[0].Equals(DBNull.Value) ? "NULL" : dr[0].ToString();
                    sa = dr[4].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[4].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    sb = dr[5].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr[5].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                    string strSQL = "insert into pub.CARDSR  VALUES  ('" + ss + "','"
                                                                         + dr[1].ToString() + "','"
                                                                         + dr[2].ToString() + "','"
                                                                         + dr[3].ToString() + "',"
                                                                         + sa + ","
                                                                         + sb + ",'"
                                                                         + dr[6].ToString() + "','"
                                                                         + dr[7].ToString() + "','"
                                                                         + dr[8].ToString() + "','"
                                                                         + dr[9].ToString() + "','"
                                                                         + dr[10].ToString() + "','"
                                                                         + dr[11].ToString() + "','"
                                                                         + dr[12].ToString() + "','"
                                                                         + dr[13].ToString() + "','"
                                                                         + dr[14].ToString() + "')";

                    cmd.CommandText = strSQL;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        WriteLog("-->230CardTimeInsertError " + strSQL + '\r', 2);
                        strInsertError += "230CardTimeInsertError "+ex.Message + '\r';
                        continue;
                    }
                    //}

                }


                tx.Commit();
                result = "OK: 230CardTimeToProgress";
                WriteLog("-->230CardTimeToProgressOK" + '\r', 1);



            }
            catch (System.Exception ex)
            {
                WriteLog("-->230CardTimeToProgressError" + ex.Message + '\r', 2);
                result = "NG: 230CardTimeToProgressError" + ex.Message;
                tx.Rollback();
            }
            finally
            {
                tx.Dispose();
                odbcCon.Close();

            }
            return result;
            #endregion


        }

        private void InsertMySql_LmtDept(string strBG)
        {
            ProgressHelp ProHelp=new ProgressHelp();
            if (strBG == "CSBG")
            {
                string SQL = "SELECT bmbh, bmmc, kbmbh, kbmmc FROM PUB.deptcsr";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog( "-->Query ProgressDeptError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog( "-->Query ProgressDeptSum:0" + '\r', 2);
                    return;
                }
                WriteLog( "-->Query ProgressDeptSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.lmt_dept";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    strSQL = "insert into swipecard.lmt_dept VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "',"
                                         + sa + ","
                                         + sb + ","
                                         + sc + "),";
                    }
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();
                    tx.Commit();
                    WriteLog( "-->insert111_DeptOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->insert111_DeptError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string SQL = "SELECT bmbh, bmmc, kbmbh, kbmmc FROM PUB.deptcsr";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog( "-->Query ProgressDeptError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog( "-->Query ProgressDeptSum:0" + '\r', 2);
                    return;
                }
                WriteLog( "-->Query ProgressdeptSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.lmt_dept";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    strSQL = "insert into swipecard.lmt_dept VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "',"
                                         + sa + ","
                                         + sb + ","
                                         + sc + "),";
                    }
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();
                    tx.Commit();
                    WriteLog( "-->insert112_DeptOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->insert112_DeptError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
 
        }

        private void InsertMysql_dept_relation(string strBG)
        {
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                string SQL = "SELECT DEPT_UP, DEPT_CD, \"LEVEL\", EXP_DEPT FROM pub.deptment2";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog("-->Query ProgressDeptMent2Error" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressDeptMent2Sum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressDeptMent2Sum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex) { WriteLog("-->ConnMysql111Error" + ex.Message + '\r', 2); return; }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.dept_relation";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    int InsertSumk = 0;
                    strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        ++InsertSumk;
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "','"
                                         + dr[1].ToString() + "','"
                                         + dr[2].ToString() + "','"
                                         + dr[3].ToString() + "'),";
                        if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                        }
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //tx.Commit();
                    }
                    tx.Commit();
                    WriteLog("-->insert111_Dept_relationOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_Dept_relationError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string SQL = "SELECT DEPT_UP, DEPT_CD, \"LEVEL\", EXP_DEPT FROM pub.deptment2";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog("-->Query ProgressDeptMent2Error" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressDeptMent2Sum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressDeptMent2Sum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex) { WriteLog("-->ConnMysql112Error" + ex.Message + '\r', 2); return; }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.dept_relation";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    int InsertSumk = 0;
                    strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        ++InsertSumk;
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "','"
                                         + dr[1].ToString() + "','"
                                         + dr[2].ToString() + "','"
                                         + dr[3].ToString() + "'),";
                        if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                        }
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //tx.Commit();
                    }
                    tx.Commit();
                    WriteLog("-->insert112_Dept_relationOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_Dept_relationError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "BBBB")
            {
                string SQL = "SELECT DEPT_UP, DEPT_CD, \"LEVEL\", EXP_DEPT FROM pub.deptment2";
                DataTable DT1 = ProHelp.QueryProgress(SQL);
                if (DT1 == null)
                {
                    WriteLog("-->Query ProgressDeptMent2Error" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressDeptMent2Sum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressDeptMent2Sum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.65.230;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex) { WriteLog("-->ConnMysql230Error" + ex.Message + '\r', 2); return; }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                try
                {
                    string strSQL = "delete from swipecard.dept_relation";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    string sa = "";
                    string sb = "";
                    string sc = "";
                    int InsertSumk = 0;
                    strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        ++InsertSumk;
                        sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + dr[1].ToString() + "'";
                        sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + dr[2].ToString() + "'";
                        sc = dr[3].Equals(DBNull.Value) ? "NULL" : "'" + dr[3].ToString() + "'";
                        strSQL += "  ('" + dr[0].ToString() + "','"
                                         + dr[1].ToString() + "','"
                                         + dr[2].ToString() + "','"
                                         + dr[3].ToString() + "'),";
                        if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.dept_relation (parent_dept,depid,dept_level,costid) VALUES ";
                        }
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //tx.Commit();
                    }
                    tx.Commit();
                    WriteLog("-->insert230_Dept_relationOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert230_Dept_relationError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
 
        }


        private void InsertMySql_empClass(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            if (strBG == "CSBG")
            {
                //60.111
                string strEMP = "";
                string strSqlMySqlEmp = "select distinct ID from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog( "-->Query 111EmpError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog( "-->Query 111EmpSUM:0" + '\r', 2);
                    return;
                }
                
                WriteLog( "-->Query 111EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);
                foreach (DataRow dr in DT1.Rows)
                {
                    strEMP += " ygbh='" + dr[0].ToString() + "' OR";
                }
                strEMP = strEMP.Substring(0, strEMP.Length - 2);


                String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog( "-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog( "-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;
 
                }
                WriteLog( "-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int i = 1;
                try
                {
                    string strSQL = "delete from swipecard.emp_class where emp_date>=curdate()-3 and emp_date<curdate()+2 and ID in (select distinct ID from swipecard.testemployee)";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    strSQL = "insert into swipecard.emp_class VALUES ";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                        strSQL += " ('" + dr[0].ToString() + "','"
                                        + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd") + "','"
                                        + sa + "',"
                                        + "curdate()" + "),";
                        if (i > 0 && i % 3000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.emp_class  VALUES ";
                        }
                        i++;
                    }
                    if (DT2.Rows.Count % 3000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //throw new ArgumentOutOfRangeException("shoudong");
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }
                    WriteLog( "-->insert111_EmpClassOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string strEMP = "";
                string strSqlMySqlEmp = "select distinct ID from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog( "-->Query 112EmpError" + '\r', 2);
                    return;
                }
                else if (DT1.Rows.Count == 0)
                {
                    WriteLog( "-->Query 112EmpSUM:0" + '\r', 2);
                    return;
                }
                WriteLog( "-->Query 112EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);
                foreach (DataRow dr in DT1.Rows)
                {
                    strEMP += " ygbh='" + dr[0].ToString() + "' OR";
                }
                strEMP = strEMP.Substring(0, strEMP.Length - 2);


                String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where ygbh!='' and ygbh is not null "; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog( "-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog( "-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                WriteLog( "-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int i = 1;
                try
                {
                    string strSQL = "delete from swipecard.emp_class where emp_date>=curdate()-3 and emp_date<curdate()+2 and ID in (select distinct ID from swipecard.testemployee)";
                    //WriteLog(DateTime.Now + "-->delete:" + " " + '\r', 1);
                    cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();

                    strSQL = "insert into swipecard.emp_class VALUES ";
                    foreach (DataRow dr in DT2.Rows)
                    {
                        string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                        strSQL += " ('" + dr[0].ToString() + "','"
                                        + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd") + "','"
                                        + sa + "',"
                                        + "curdate()" + "),";
                        if (i > 0 && i % 3000 == 0)
                        {
                            strSQL = strSQL.Substring(0, strSQL.Length - 1);
                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                            strSQL = "insert into swipecard.emp_class  VALUES ";
                        }
                        i++;
                    }
                    if (DT2.Rows.Count % 3000 != 0)
                    {
                        strSQL = strSQL.Substring(0, strSQL.Length - 1);
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                        //throw new ArgumentOutOfRangeException("shoudong");
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }
                    WriteLog( "-->insert112_EmpClassOK:" + " " + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog( "-->insert111_EmpClassOK,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
               
 
            }
        }

        private void InsertMySql_empClass2(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111
                string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr  in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
               

                WriteLog("-->Query 111EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String sql = "SELECT ygbh,fdate,bc FROM KQ_RECORD_BC WHERE ygbh IN (SELECT DISTINCT 編號  FROM V_RYBZ_TX ) AND fdate>=CONVERT(varchar(100), GETDATE(), 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) ORDER BY ygbh, fdate"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                
                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //如果工号,日期,班别都存在，这不用处理。如果不存在则继续判断
                        if (!ClassArray3.Contains(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString()))
                        {
                            //如果工号和日期都存在，则Update班别。
                            if (ClassArray2.Contains(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                            {
                                string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                                strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                                //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                            }
                            else
                            {
                                //如果工号和日期都不存在，则插入
                                ++InsertSumk;
                                string sb = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                                strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                     + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                     + sb + "',"
                                                     + "curdate()" + "),";
                                
                                if (InsertSumk > 0 && InsertSumk % 5000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                }
                                
                            }
                        }
                    }
                    if (InsertSumk % 5000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112
                string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }


                WriteLog("-->Query 111EmpSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                String sql = "SELECT ygbh,fdate,bc FROM KQ_RECORD_BC WHERE ygbh IN (SELECT DISTINCT 編號  FROM V_RYBZ_LZJ ) AND fdate>=CONVERT(varchar(100), GETDATE(), 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) ORDER BY ygbh, fdate"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB112(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT2.Rows)
                    {
                        //如果工号,日期,班别都存在，这不用处理。如果不存在则继续判断
                        if (!ClassArray3.Contains(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString()))
                        {
                            //如果工号和日期都存在，则Update班别。
                            if (ClassArray2.Contains(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                            {
                                string sa = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                                strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                                //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                            }
                            else
                            {
                                //如果工号和日期都不存在，则插入.
                                ++InsertSumk;
                                string sb = dr[2].ToString() == "\\" ? "\\\\" : dr[2].ToString();
                                strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                     + Convert.ToDateTime(dr[1].ToString()).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                     + sb + "',"
                                                     + "curdate()" + "),";

                                if (InsertSumk > 0 && InsertSumk % 5000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                }

                            }
                        }
                    }
                    if (InsertSumk % 5000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            
        }

        private void InsertMySql_empClass1530(string strBG)
        {
            //如果Mysql有当天资料但是班别为空，则update，如果Mysql不存在当天资料则Insert
            //Update和Insert明天资料。
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
            //string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where  emp_date=date_add(CURDATE(),interval 1 day) order by ID, emp_date";
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111

                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 111ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 30;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //WriteLog(dr[0].ToString() + '\r', 1);
                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    if (i == 0)
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where (class_no ='' or class_no='39') and ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    else
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    cmd.CommandText = strSQLUpdate;
                                    int it = cmd.ExecuteNonQuery();
                                    if (it > 0) UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {

                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                       
                        //tx.Commit();
                    }
                    //else
                    //{
                        
                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別39
                    strSQLUpdate = "update swipecard.emp_class set class_no='39' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();

                    
                    tx.Commit();

                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112

                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 112ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {

                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    if (i == 0)
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where (class_no ='' or class_no='39') and ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    else
                                    { strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'"; }
                                    cmd.CommandText = strSQLUpdate;
                                    int it = cmd.ExecuteNonQuery();
                                    if (it > 0) UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {
                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        
                        //tx.Commit();
                    }
                    //else
                    //{
                       
                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別39 2017-10-18
                    strSQLUpdate = "update swipecard.emp_class set class_no='39' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();

                    
                    tx.Commit();

                    WriteLog("-->Update112_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
 
        }

        private void InsertMySql_empClass0330(string strBG)
        {
            //今天和明天的都会Update和Insert
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) order by ID, emp_date";
            string strSqlEmp = "select id from swipecard.testemployee where isonwork=0";
            //string strSqlMySqlEmp = "select ID, emp_date, class_no FROM swipecard.emp_class where  emp_date=date_add(CURDATE(),interval 1 day) order by ID, emp_date";
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111

                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111ClassError" + '\r', 2);
                    return;
                }

                DataTable DT2 = MySqlHelp.QueryDBC(strSqlEmp);
                if (DT2 == null)
                {
                    WriteLog("-->Query 111EmpError" + '\r', 2);
                    return;
                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["ID"] };//设置第一列为主键 

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 111ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 30;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //如果Mysql用户表中无此用户，继续循环下一个。
                        DataRow Tempdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        if (Tempdr == null) { continue; }

                        //WriteLog(dr[0].ToString() + '\r', 1);
                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue; 
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {

                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        //tx.Commit();
                    }
                    //else
                    //{
                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別39
                    strSQLUpdate = "update swipecard.emp_class set class_no='39' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    tx.Commit();

                    
                    

                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112

                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 112ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        
                            TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                            int d = ts.Days;
                            string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                            List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                            for (int i = 0; i < 2; i++)
                            {
                                if ((d + i) >= list.Count) continue; 
                                if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                                {
                                    if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                    {
                                        string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                        strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                        cmd.CommandText = strSQLUpdate;
                                        cmd.ExecuteNonQuery();
                                        UpdateSumk++;
                                        //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                    }
                                    else
                                    {
                                        ++InsertSumk;
                                        string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                        strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                             + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                             + sb + "',"
                                                             + "curdate()" + "),";

                                        if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                        {
                                            strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                            cmd.CommandText = strSQLInsert;
                                            cmd.ExecuteNonQuery();

                                            strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                        }

                                    }
                                }
                            }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                    
                        //tx.Commit();
                    }
                    //else
                    //{
                       
                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    //如果emp_class有資料，但是班別為空，則給默認班別39
                    strSQLUpdate = "update swipecard.emp_class set class_no='39' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    tx.Commit();

                    

                    WriteLog("-->Update112_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmpClassOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            else if (strBG == "JF")
            {
                #region JF
                //10.64.155.200

                DataTable DT1 = MySqlHelp.Query155_200(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112ClassError" + '\r', 2);
                    return;
                }

                List<string> ClassArray3 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray3.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + dr[2].ToString());
                }
                List<string> ClassArray2 = new List<string>();
                foreach (DataRow dr in DT1.Rows)
                {
                    ClassArray2.Add(dr[0].ToString() + Convert.ToDateTime(dr[1].ToString()).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo));
                }
                WriteLog("-->Query 112ClassSum:" + DT1.Rows.Count.ToString() + '\r', 1);

                //String sql = "SELECT DISTINCT 編號  FROM V_RYBZ_TX"; // 查询语句
                ////String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                //DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                //if (DT2 == null)
                //{
                //    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                //    return;
                //}
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;
                //}
                //string strEMP = "";
                //foreach (DataRow dr in DT2.Rows)
                //{
                //    strEMP += " EMP_CD='" + dr[0].ToString() + "' OR";
                //}
                //strEMP = strEMP.Substring(0, strEMP.Length - 2);
                //WriteLog("-->SelectSqlServerEmpSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, START_DATE, CLASS_CD from PUB.EMPCLASSR ";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEmpClassSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEmpClassSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEmpClassSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.emp_class VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {

                        TimeSpan ts = DateTime.Now.Date - Convert.ToDateTime(dr[1].ToString());
                        int d = ts.Days;
                        string strClass = dr[2].ToString().Substring(0, dr[2].ToString().Length - 1);
                        List<string> list = strClass.Split(new string[] { "," }, StringSplitOptions.None).ToList<string>();
                        for (int i = 0; i < 2; i++)
                        {
                            if ((d + i) >= list.Count) continue;
                            if (!ClassArray3.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + list[d + i]))
                            {
                                if (ClassArray2.Contains(dr[0].ToString() + DateTime.Now.AddDays(i).ToString("yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)))
                                {
                                    string sa = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLUpdate = "update swipecard.emp_class SET class_no ='" + sa + "',update_time=curdate() where ID='" + dr[0].ToString() + "'and emp_date='" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "'";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpClassOK:" + k + strSQLUpdate + '\r', 1);
                                }
                                else
                                {
                                    ++InsertSumk;
                                    string sb = list[d + i] == "\\" ? "\\\\" : list[d + i];
                                    strSQLInsert += " ('" + dr[0].ToString() + "','"
                                                         + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','"
                                                         + sb + "',"
                                                         + "curdate()" + "),";

                                    if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                                    {
                                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                        cmd.CommandText = strSQLInsert;
                                        cmd.ExecuteNonQuery();

                                        strSQLInsert = "insert into swipecard.emp_class  VALUES ";
                                    }

                                }
                            }
                        }

                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();

                        //tx.Commit();
                    }
                    //else
                    //{

                    //    tx.Commit();
                    //}
                    //如果emp_class有資料，但是班別為空，則給默認班別4
                    strSQLUpdate = "update swipecard.emp_class set class_no='4' where emp_date>=curdate() and emp_date<date_add(CURDATE(),interval 2 day) and class_no=''";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    tx.Commit();
                    WriteLog("-->Update112_EmpClassOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmpClassOK:" + InsertSumk + '\r', 1);

                    
                    
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }

        }

        private void InsertMysql_empClass4(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 111DBError" + ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                //int UpdateSumk = 0;
                //int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";
                try
                {
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為39
                    strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate())";
                    DataTable DTID = MySqlHelp.QueryDBC(strSQLUpdate);
                    if (DTID == null)
                    {
                        WriteLog("-->Query 111TestemployeeError" + '\r', 2);
                        return;
                    }
                    foreach (DataRow dr in DTID.Rows)
                    {
                        strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate(),'39',curdate())";
                        cmd.CommandText = strSQLUpdate;
                        cmd.ExecuteNonQuery();
                        Insert4++;
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClass4:" + Insert4 + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClass4Error,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 112DBError"+ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                //int UpdateSumk = 0;
                //int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";
                try
                {
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
                    //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為39
                    strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate())";
                    DataTable DTID = MySqlHelp.QueryDbAsbg(strSQLUpdate);
                    if (DTID == null)
                    {
                        WriteLog("-->Query 112TestemployeeError" + '\r', 2);
                        return;
                    }
                    foreach (DataRow dr in DTID.Rows)
                    {
                        strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate(),'39',curdate())";
                        cmd.CommandText = strSQLUpdate;
                        cmd.ExecuteNonQuery();
                        Insert4++;
                    }
                    tx.Commit();
                    WriteLog("-->Update112_EmpClass4:" + Insert4 + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClass4Error,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "JF")
            {
                String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                //int UpdateSumk = 0;
                //int InsertSumk = 0;
                int Insert4 = 0;
                string strSQLUpdate = "";

                //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
                strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate())";
                DataTable DTID = MySqlHelp.Query155_200(strSQLUpdate);
                foreach (DataRow dr in DTID.Rows)
                {
                    strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate(),'4',curdate())";
                    cmd.CommandText = strSQLUpdate;
                    cmd.ExecuteNonQuery();
                    Insert4++;
                }
                tx.Commit();
                WriteLog("-->Update155_EmpClass4:" + Insert4 + '\r', 1);
            }
        }

        private void InsertMySql_testemployee(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                //60.111
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.Query155_200(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }
                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                //String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                        //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                        if (Tempdr != null)
                        {
                            if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != dr["cardid"].ToString() ||
                                Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() == dr["EMP_STATUS"].ToString())
                            {
                                if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                { Istatus = 1; }
                                else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                { Istatus = 0; }
                                strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                              "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                              "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                              "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                              "',cardid = '" + dr["cardid"].ToString() +
                                                                              "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                              "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                            }
                        }
                        else
                        {
                            //如果Tempder为空，说明Mysql中要Insert
                            ++InsertSumk;
                            strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                  + dr["EMP_NAME"].ToString() + "','"
                                                  + dr["DEPT_CD"].ToString() + "','"
                                                  + dr["DEPT_NAME"].ToString() + "','"
                                                  + dr["D_I_CD"].ToString() + "','"
                                                  + dr["cardid"].ToString() + "','"
                                                  + dr["EXP_DEPT"].ToString() + "',"
                                                  + "curdate()" + "),";

                            if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                            {
                                strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                cmd.CommandText = strSQLInsert;
                                cmd.ExecuteNonQuery();

                                strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                            }
                        }
                    }
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
               
            }
 
        }

        private void InsertMySql_testemployee_SqlServer(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT [編號] id, [姓名] name, [部門編號] depid, [部門名稱] depname, [直間接] direct, [卡號] cardid, [課部門編號] costid FROM V_RYBZ_TX WHERE [更新標誌]='N'"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["id"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                //DataTable DT3 = ProHelp.QueryProgress(SQL);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                //    return;
                //}
                //else if (DT3.Rows.Count == 0)
                //{
                //    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                //    return;
                //}
                
                //WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 111DBError" + ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLUpdateIswork = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    //foreach (DataRow dr in DT1.Rows)
                    //{
                    //    DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                    //    if (MSdr == null)
                    //    {
                    //        strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                    //        cmd.CommandText = strSQLUpdateIswork;
                    //        int i = cmd.ExecuteNonQuery();
                    //        if (i > 0) { WriteLog("-->Update111_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                    //    }
                        
                    //}
                    foreach (DataRow dr in DT2.Rows)
                    {

                            DataRow Tempdr = DT1.Rows.Find(dr["id"].ToString());
                            //如果Tempdr为空说明Mysql中没有此员工，需要insert
                            //如果Tempdr不为空说明Mysql中存在此员工，继续判断是要Update还是Insert
                            if (Tempdr != null)
                            {
                                if (Convert.ToInt32(Tempdr["isOnWork"].ToString()) == 0)
                                {
                                    if (Tempdr["ID"].ToString() != dr["id"].ToString() || Tempdr["Name"].ToString() != dr["name"].ToString() ||
                                        Tempdr["depid"].ToString() != dr["depid"].ToString() || Tempdr["depname"].ToString() != dr["depname"].ToString() ||
                                        Tempdr["Direct"].ToString() != dr["direct"].ToString() || Tempdr["cardid"].ToString() != dr["cardid"].ToString() ||
                                        Tempdr["costID"].ToString() != dr["costid"].ToString())
                                    {
                                        //if (dr["EMP_STATUS"].ToString() == string.Empty)
                                        //{ Istatus = 0; }
                                        //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                        //{ Istatus = 1; }
                                        //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                        //{ Istatus = 0; }
                                        strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["name"].ToString() +
                                                                                      "',depid = '" + dr["depid"].ToString() +
                                                                                      "',depname = '" + dr["depname"].ToString() +
                                                                                      "',Direct = '" + dr["direct"].ToString() +
                                                                                      "',cardid = '" + dr["cardid"].ToString() +
                                                                                      "',costID = '" + dr["costid"].ToString() +
                                                                                      "',updateDate =curdate() WHERE ID = '" + dr["id"].ToString() + "' ";
                                        cmd.CommandText = strSQLUpdate;
                                        cmd.ExecuteNonQuery();
                                        UpdateSumk++;
                                        //WriteLog("-->Update111_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                    }
                                }
                            }
                            else
                            {
                                //如果Tempder为空，说明Mysql中要Insert
                                ++InsertSumk;
                                strSQLInsert += " ('" + dr["id"].ToString() + "','"
                                                      + dr["name"].ToString() + "','"
                                                      + dr["depid"].ToString() + "','"
                                                      + dr["depname"].ToString() + "','"
                                                      + dr["direct"].ToString() + "','"
                                                      + dr["cardid"].ToString() + "','"
                                                      + dr["costid"].ToString() + "',"
                                                      + "curdate()" + "),";
                                //WriteLog("-->Insert111_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                                //                      + dr["EMP_NAME"].ToString() + "','"
                                //                      + dr["DEPT_CD"].ToString() + "','"
                                //                      + dr["DEPT_NAME"].ToString() + "','"
                                //                      + dr["D_I_CD"].ToString() + "','"
                                //                      + SqlServerdr["kh"].ToString() + "','"
                                //                      + dr["EXP_DEPT"].ToString() + "',"
                                //                      + "curdate()" + '\r', 1);
                                if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                                }
                            }
                        
                    }
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion

            }
            else if (strBG == "ASBG")
            {
                //60.112
                #region ASBG
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT [編號] id, [姓名] name, [部門編號] depid, [部門名稱] depname, [直間接] direct, [卡號] cardid, [課部門編號] costid FROM V_RYBZ_LZJ WHERE [更新標誌]='N'"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB112(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["id"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                //string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                //DataTable DT3 = ProHelp.QueryProgress(SQL);
                //if (DT3 == null)
                //{
                //    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                //    return;
                //}
                //else if (DT3.Rows.Count == 0)
                //{
                //    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                //    return;
                //}

                //WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 112DBError" + ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLUpdateIswork = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    //foreach (DataRow dr in DT1.Rows)
                    //{
                    //    DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                    //    if (MSdr == null)
                    //    {
                    //        strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                    //        cmd.CommandText = strSQLUpdateIswork;
                    //        int i = cmd.ExecuteNonQuery();
                    //        if (i > 0) { WriteLog("-->Update111_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                    //    }

                    //}
                    foreach (DataRow dr in DT2.Rows)
                    {

                        DataRow Tempdr = DT1.Rows.Find(dr["id"].ToString());
                        //如果Tempdr为空说明Mysql中没有此员工，需要insert
                        //如果Tempdr不为空说明Mysql中存在此员工，继续判断是要Update还是Insert
                        if (Tempdr != null)
                        {
                            if (Convert.ToInt32(Tempdr["isOnWork"].ToString()) == 0)
                            {
                                if (Tempdr["ID"].ToString() != dr["id"].ToString() || Tempdr["Name"].ToString() != dr["name"].ToString() ||
                                    Tempdr["depid"].ToString() != dr["depid"].ToString() || Tempdr["depname"].ToString() != dr["depname"].ToString() ||
                                    Tempdr["Direct"].ToString() != dr["direct"].ToString() || Tempdr["cardid"].ToString() != dr["cardid"].ToString() ||
                                    Tempdr["costID"].ToString() != dr["costid"].ToString())
                                {
                                    //if (dr["EMP_STATUS"].ToString() == string.Empty)
                                    //{ Istatus = 0; }
                                    //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                    //{ Istatus = 1; }
                                    //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                    //{ Istatus = 0; }
                                    strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["name"].ToString() +
                                                                                  "',depid = '" + dr["depid"].ToString() +
                                                                                  "',depname = '" + dr["depname"].ToString() +
                                                                                  "',Direct = '" + dr["direct"].ToString() +
                                                                                  "',cardid = '" + dr["cardid"].ToString() +
                                                                                  "',costID = '" + dr["costid"].ToString() +
                                                                                  "',updateDate =curdate() WHERE ID = '" + dr["id"].ToString() + "' ";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                }
                            }
                        }
                        else
                        {
                            //如果Tempder为空，说明Mysql中要Insert
                            ++InsertSumk;
                            strSQLInsert += " ('" + dr["id"].ToString() + "','"
                                                  + dr["name"].ToString() + "','"
                                                  + dr["depid"].ToString() + "','"
                                                  + dr["depname"].ToString() + "','"
                                                  + dr["direct"].ToString() + "','"
                                                  + dr["cardid"].ToString() + "','"
                                                  + dr["costid"].ToString() + "',"
                                                  + "curdate()" + "),";
                            //WriteLog("-->Insert111_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                            //                      + dr["EMP_NAME"].ToString() + "','"
                            //                      + dr["DEPT_CD"].ToString() + "','"
                            //                      + dr["DEPT_NAME"].ToString() + "','"
                            //                      + dr["D_I_CD"].ToString() + "','"
                            //                      + SqlServerdr["kh"].ToString() + "','"
                            //                      + dr["EXP_DEPT"].ToString() + "',"
                            //                      + "curdate()" + '\r', 1);
                            if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                            {
                                strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                cmd.CommandText = strSQLInsert;
                                cmd.ExecuteNonQuery();

                                strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                            }
                        }

                    }
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update112_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion

            }

        }

        private void InsertMySql_testemployee_Progress2(string strBG)
        {
            //增加
            //如果新員工在emp_class中沒有資料，則要插入兩筆資料，今天和明天的班別資料，並且默認為4班別。
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            if (strBG == "CSBG")
            {
                //60.111
                string sTempID = "";
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_TX"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["bh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD, CRN, EXP_DEPT, EMP_STATUS FROM pub.EMPR WHERE SUBSTR(EXP_DEPT_NAME,1,2)='通訊'";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }

                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLUpdateIswork = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                        if (MSdr == null)
                        {
                            strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                            cmd.CommandText = strSQLUpdateIswork;
                            int i = cmd.ExecuteNonQuery();
                            if (i > 0) { WriteLog("-->Update111_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                        }

                    }
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        //DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr为空说明Sqlserver中不存在此员工，不做处理。。
                        //if (SqlServerdr != null)
                        //{
                            DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                            //如果SqlServerdr不为空说明sqlserver中存在此员工，继续判断是要Update还是Insert
                            //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                            if (Tempdr != null)
                            {
                                if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                    Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                    Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || 
                                    Tempdr["isOnWork"].ToString() == dr["EMP_STATUS"].ToString())
                                {
                                    if (dr["EMP_STATUS"].ToString() == string.Empty)
                                    { Istatus = 0; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                    { Istatus = 1; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                    { Istatus = 0; }
                                    strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                                  "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                                  "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                                  "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                                  //"',cardid = '" + SqlServerdr["kh"].ToString() +
                                                                                  "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                                  "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update111_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                }
                            }
                            else
                            {
                                //如果Tempder为空，说明Mysql中要Insert
                                sTempID += " emp.ID='" + dr["EMP_CD"].ToString() + "' OR";
                                ++InsertSumk;
                                strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                      + dr["EMP_NAME"].ToString() + "','"
                                                      + dr["DEPT_CD"].ToString() + "','"
                                                      + dr["DEPT_NAME"].ToString() + "','"
                                                      + dr["D_I_CD"].ToString() + "','"
                                                      //+ SqlServerdr["kh"].ToString() + "','"
                                                      + dr["EXP_DEPT"].ToString() + "',"
                                                      + "curdate()" + "),";
                                //WriteLog("-->Insert111_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                                //                      + dr["EMP_NAME"].ToString() + "','"
                                //                      + dr["DEPT_CD"].ToString() + "','"
                                //                      + dr["DEPT_NAME"].ToString() + "','"
                                //                      + dr["D_I_CD"].ToString() + "','"
                                //                      + SqlServerdr["kh"].ToString() + "','"
                                //                      + dr["EXP_DEPT"].ToString() + "',"
                                //                      + "curdate()" + '\r', 1);
                                if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                                }
                            }
                        //}
                    }
                    sTempID = sTempID.Substring(0, sTempID.Length - 2);
                    for (int i = 0; i < 2; i++)
                    {
                        strSQLUpdate = @"SELECT b.ID
  FROM (SELECT emp.ID
          FROM swipecard.testemployee emp
         WHERE (" + sTempID + @")) b
 WHERE b.ID NOT IN
          (SELECT empclass.ID
             FROM swipecard.emp_class empclass
            WHERE     emp_date = '" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' AND (" + sTempID + "))";
                        DataTable DTID = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                        foreach (DataRow dr in DTID.Rows)
                        {
                            strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "','"
                                                                                   + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','4',curdate())";
                            cmd.CommandText = strSQLUpdate;
                            cmd.ExecuteNonQuery();
                        }
                    }                                                             
                        
                    
                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }

            }
            else if (strBG == "ASBG")
            {
                //60.112
                string sTempID = "";
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_LZJ"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB112(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                else if (DT2.Rows.Count == 0)
                {
                    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                    return;

                }
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["bh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD,'cardid', EXP_DEPT, EMP_STATUS FROM PUB.EMPR";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }

                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                //String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    string strSQLUpdateIswork = "";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        DataRow MSdr = DT2.Rows.Find(dr["ID"].ToString());
                        if (MSdr == null)
                        {
                            strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                            cmd.CommandText = strSQLUpdateIswork;
                            int i = cmd.ExecuteNonQuery();
                            if (i > 0) { WriteLog("-->Update112_" + dr["ID"].ToString() + " Iswork=1OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                        }

                    }
                    foreach (DataRow dr in DT3.Rows)
                    {
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr为空说明Sqlserver中不存在此员工，不做处理。
                        if (SqlServerdr != null)
                        {
                            DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                            //如果SqlServerdr不为空说明sqlserver中存在此员工，继续判断是要Update还是Insert
                            //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                            if (Tempdr != null)
                            {
                                if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                    Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                    Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != SqlServerdr["kh"].ToString() ||
                                    Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() == dr["EMP_STATUS"].ToString())
                                {
                                    if (dr["EMP_STATUS"].ToString() == string.Empty)
                                    { Istatus = 0; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                    { Istatus = 1; }
                                    else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                    { Istatus = 0; }
                                    strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                                  "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                                  "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                                  "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                                  "',cardid = '" + SqlServerdr["kh"].ToString() +
                                                                                  "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                                  "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                    cmd.CommandText = strSQLUpdate;
                                    cmd.ExecuteNonQuery();
                                    UpdateSumk++;
                                    //WriteLog("-->Update112_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                                }
                            }
                            else
                            {
                                //如果Tempder为空，说明Mysql中要Insert
                                sTempID += " emp.ID='" + dr["EMP_CD"].ToString() + "' OR";
                                ++InsertSumk;
                                strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                      + dr["EMP_NAME"].ToString() + "','"
                                                      + dr["DEPT_CD"].ToString() + "','"
                                                      + dr["DEPT_NAME"].ToString() + "','"
                                                      + dr["D_I_CD"].ToString() + "','"
                                                      + SqlServerdr["kh"].ToString() + "','"
                                                      + dr["EXP_DEPT"].ToString() + "',"
                                                      + "curdate()" + "),";
                                //WriteLog("-->Insert112_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                                //                      + dr["EMP_NAME"].ToString() + "','"
                                //                      + dr["DEPT_CD"].ToString() + "','"
                                //                      + dr["DEPT_NAME"].ToString() + "','"
                                //                      + dr["D_I_CD"].ToString() + "','"
                                //                      + SqlServerdr["kh"].ToString() + "','"
                                //                      + dr["EXP_DEPT"].ToString() + "',"
                                //                      + "curdate()" + '\r', 1);

                                if (InsertSumk > 0 && InsertSumk % 1000 == 0)
                                {
                                    strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                    cmd.CommandText = strSQLInsert;
                                    cmd.ExecuteNonQuery();

                                    strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                                }
                            }
                        }
                    }
                    sTempID = sTempID.Substring(0, sTempID.Length - 2);
                    for (int i = 0; i < 2; i++)
                    {
                        strSQLUpdate = @"SELECT b.ID
  FROM (SELECT emp.ID
          FROM swipecard.testemployee emp
         WHERE (" + sTempID + @")) b
 WHERE b.ID NOT IN
          (SELECT empclass.ID
             FROM swipecard.emp_class empclass
            WHERE     emp_date = '" + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' AND (" + sTempID + "))";
                        DataTable DTID = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                        foreach (DataRow dr in DTID.Rows)
                        {
                            strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "','"
                                                                                   + DateTime.Now.AddDays(i).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "','4',curdate())";
                            cmd.CommandText = strSQLUpdate;
                            cmd.ExecuteNonQuery();
                        }
                    } 

                    if (InsertSumk % 1000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update112_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }

            }

        }

        private void InsertMySql_testemployee_Progress3(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();
            
            if (strBG == "CSBG")
            {
                #region CSBG
                //60.111
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 111TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_TX"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;

                //}
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["bh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD, CRN, EXP_DEPT, EMP_STATUS FROM pub.EMPR WHERE SUBSTR(EXP_DEPT_NAME,1,2)='通訊'";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }
                DT3.PrimaryKey = new DataColumn[] { DT3.Columns["EMP_CD"] };//设置第一列为主键 
                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 111DBError" + ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLUpdateIswork = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    //foreach (DataRow dr in DT1.Rows)
                    //{
                    //    DataRow MSdr = DT3.Rows.Find(dr["ID"].ToString());
                    //    if (MSdr == null)
                    //    {
                    //        strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                    //        cmd.CommandText = strSQLUpdateIswork;
                    //        cmd.ExecuteNonQuery();
                    //        //int i = cmd.ExecuteNonQuery();
                    //        //if (i > 0) { WriteLog("-->Update111_" + dr["ID"].ToString() + " Iswork=1 OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                    //    }

                    //}
                    foreach (DataRow dr in DT3.Rows)
                    {
                        string strCardID = "";
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr为空说明Sqlserver中不存在此员工，cardID设置为空,不为空则把sqlserver中的cardID给strCardID
                        if (SqlServerdr != null)
                        {
                            if (SqlServerdr["kh"].ToString() != string.Empty)
                            {
                                strCardID = SqlServerdr["kh"].ToString();
                            }
                            else
                            {
                                strCardID = dr["CRN"].ToString();
                            }
                        }
                        else
                        {
                            strCardID = dr["CRN"].ToString();
 
                        }

                        DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr不为空说明sqlserver中存在此员工，继续判断是要Update还是Insert
                        //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                        if (dr["EMP_STATUS"].ToString() == string.Empty)
                        { Istatus = 0; }
                        else
                        {
                            switch (Convert.ToInt32(dr["EMP_STATUS"].ToString()))
                            {
                                case 1:
                                case 8:
                                    Istatus = 0;
                                    break;
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                case 6:
                                case 7:
                                    Istatus = 1;
                                    break;
                                default:
                                    Istatus = 0;
                                    break;
                            }
                        }
                        if (Tempdr != null)
                        {
                            if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != strCardID ||
                                Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() != Istatus.ToString())
                            {
                                //員工狀態  "1"是在職 "2"停薪 "3" 離職 4 "自離  "5 開除 6  資遣7 試用不合格 8 廠區異動
                                //if (dr["EMP_STATUS"].ToString() == string.Empty)
                                //{ Istatus = 0; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                //{ Istatus = 1; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                //{ Istatus = 0; }

                                strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                              "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                              "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                              "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                              "',cardid = '" + strCardID +
                                                                              "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                              "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                                //WriteLog("-->Update111_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                            }
                        }
                        else
                        {
                            //如果Tempder为空，说明Mysql中要Insert
                            ++InsertSumk;
                            strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                  + dr["EMP_NAME"].ToString() + "','"
                                                  + dr["DEPT_CD"].ToString() + "','"
                                                  + dr["DEPT_NAME"].ToString() + "','"
                                                  + dr["D_I_CD"].ToString() + "','"
                                                  + strCardID + "','"
                                                  + dr["EXP_DEPT"].ToString() + "',"
                                                  + "curdate()" + "),";
                            //WriteLog("-->Insert111_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                            //                      + dr["EMP_NAME"].ToString() + "','"
                            //                      + dr["DEPT_CD"].ToString() + "','"
                            //                      + dr["DEPT_NAME"].ToString() + "','"
                            //                      + dr["D_I_CD"].ToString() + "','"
                            //                      + SqlServerdr["kh"].ToString() + "','"
                            //                      + dr["EXP_DEPT"].ToString() + "',"
                            //                      + "curdate()" + '\r', 1);
                            if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                            {
                                strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                cmd.CommandText = strSQLInsert;
                                cmd.ExecuteNonQuery();

                                strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                            }
                        }
                        //}
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update111_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert111_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            
            else if (strBG == "ASBG")
            {
                #region ASBG
                //60.112
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_LZJ"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB112(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;

                //}
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["bh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD, CRN, EXP_DEPT, EMP_STATUS FROM pub.EMPR WHERE SUBSTR(EXP_DEPT_NAME,1,3)='零組件'";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }
                DT3.PrimaryKey = new DataColumn[] { DT3.Columns["EMP_CD"] };//设置第一列为主键
                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                //String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                try
                {
                    MySqlConn.Open();
                }
                catch (Exception ex)
                {
                    WriteLog("-->Query 112DBError" + ex.Message + '\r', 2);
                    return;
                }
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    string strSQLUpdateIswork = "";
                    //foreach (DataRow dr in DT1.Rows)
                    //{
                    //    DataRow MSdr = DT3.Rows.Find(dr["ID"].ToString());
                    //    if (MSdr == null)
                    //    {
                    //        strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                    //        cmd.CommandText = strSQLUpdateIswork;
                    //        cmd.ExecuteNonQuery();
                    //        //int i = cmd.ExecuteNonQuery();
                    //        //if (i > 0) { WriteLog("-->Update112_" + dr["ID"].ToString() + " Iswork=1 OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                    //    }

                    //}
                    foreach (DataRow dr in DT3.Rows)
                    {
                        string strCardID = "";
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr为空说明Sqlserver中不存在此员工，cardID设置为空,不为空则把sqlserver中的cardID给strCardID
                        if (SqlServerdr != null)
                        {
                            if (SqlServerdr["kh"].ToString() != string.Empty)
                            {
                                strCardID = SqlServerdr["kh"].ToString();
                            }
                            else
                            {
                                strCardID = dr["CRN"].ToString();
                            }
                        }
                        else
                        {
                            strCardID = dr["CRN"].ToString();

                        }
                        DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr不为空说明sqlserver中存在此员工，继续判断是要Update还是Insert
                        //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                        if (dr["EMP_STATUS"].ToString() == string.Empty)
                        { Istatus = 0; }
                        else
                        {
                            switch (Convert.ToInt32(dr["EMP_STATUS"].ToString()))
                            {
                                case 1:
                                case 8:
                                    Istatus = 0;
                                    break;
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                case 6:
                                case 7:
                                    Istatus = 1;
                                    break;
                                default:
                                    Istatus = 0;
                                    break;

                            }
                        }
                        if (Tempdr != null)
                        {
                            if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != strCardID ||
                                Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() != Istatus.ToString())
                            {
                                //if (dr["EMP_STATUS"].ToString() == string.Empty)
                                //{ Istatus = 0; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                //{ Istatus = 1; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                //{ Istatus = 0; }
                                strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                              "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                              "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                              "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                              "',cardid = '" + strCardID +
                                                                              "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                              "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                                //WriteLog("-->Update112_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                            }
                        }
                        else
                        {
                            //如果Tempder为空，说明Mysql中要Insert
                            ++InsertSumk;
                            strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                  + dr["EMP_NAME"].ToString() + "','"
                                                  + dr["DEPT_CD"].ToString() + "','"
                                                  + dr["DEPT_NAME"].ToString() + "','"
                                                  + dr["D_I_CD"].ToString() + "','"
                                                  + strCardID + "','"
                                                  + dr["EXP_DEPT"].ToString() + "',"
                                                  + "curdate()" + "),";
                            //WriteLog("-->Insert112_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                            //                      + dr["EMP_NAME"].ToString() + "','"
                            //                      + dr["DEPT_CD"].ToString() + "','"
                            //                      + dr["DEPT_NAME"].ToString() + "','"
                            //                      + dr["D_I_CD"].ToString() + "','"
                            //                      + SqlServerdr["kh"].ToString() + "','"
                            //                      + dr["EXP_DEPT"].ToString() + "',"
                            //                      + "curdate()" + '\r', 1);

                            if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                            {
                                strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                cmd.CommandText = strSQLInsert;
                                cmd.ExecuteNonQuery();

                                strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                            }
                        }
                        //}
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update112_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            
            else if (strBG == "JF")
            {
                #region JF
                //60.112
                string strSqlMySqlEmp = "select  ID, Name, depid, depname, Direct, cardid, costID, isOnWork from swipecard.testemployee";
                DataTable DT1 = MySqlHelp.Query155_200(strSqlMySqlEmp);
                if (DT1 == null)
                {
                    WriteLog("-->Query 112TestemployeeError" + '\r', 2);
                    return;
                }
                DT1.PrimaryKey = new DataColumn[] { DT1.Columns["ID"] };//设置第一列为主键 

                String sql = "SELECT 編號 bh, 卡號 kh FROM V_RYBZ_TX"; // 查询语句
                //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
                DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
                if (DT2 == null)
                {
                    WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                    return;
                }
                //else if (DT2.Rows.Count == 0)
                //{
                //    WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                //    return;

                //}
                DT2.PrimaryKey = new DataColumn[] { DT2.Columns["bh"] };

                WriteLog("-->SelectSqlServerSum:" + DT2.Rows.Count.ToString() + '\r', 1);

                string SQL = "SELECT EMP_CD, EMP_NAME, DEPT_CD, DEPT_NAME, D_I_CD, CRN, EXP_DEPT, EMP_STATUS FROM pub.EMPR WHERE SUBSTR(EXP_DEPT_NAME,1,2)='通訊'";
                DataTable DT3 = ProHelp.QueryProgress(SQL);
                if (DT3 == null)
                {
                    WriteLog("-->Query ProgressEMPRSumError" + '\r', 2);
                    return;
                }
                else if (DT3.Rows.Count == 0)
                {
                    WriteLog("-->Query ProgressEMPRSum:0" + '\r', 2);
                    return;
                }
                DT3.PrimaryKey = new DataColumn[] { DT3.Columns["EMP_CD"] };//设置第一列为主键
                WriteLog("-->Query ProgressEMPRSum:" + DT3.Rows.Count.ToString() + '\r', 1);

                //String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                int Istatus = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate) VALUES ";
                try
                {
                    string strSQLUpdateIswork = "";
                    foreach (DataRow dr in DT1.Rows)
                    {
                        DataRow MSdr = DT3.Rows.Find(dr["ID"].ToString());
                        if (MSdr == null)
                        {
                            strSQLUpdateIswork = "update swipecard.testemployee SET isOnWork =1,updateDate =curdate() where isOnWork=0 and ID = '" + dr["ID"].ToString() + "'";
                            cmd.CommandText = strSQLUpdateIswork;
                            cmd.ExecuteNonQuery();
                            //int i = cmd.ExecuteNonQuery();
                            //if (i > 0) { WriteLog("-->Update112_" + dr["ID"].ToString() + " Iswork=1 OK:" + UpdateSumk + strSQLUpdate + '\r', 1); }
                        }

                    }
                    foreach (DataRow dr in DT3.Rows)
                    {
                        string strCardID = "";
                        //DataRow[] Tempdr = DT1.Select("ID='" + dr["EMP_CD"].ToString() + "'");
                        DataRow SqlServerdr = DT2.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr为空说明Sqlserver中不存在此员工，cardID设置为空,不为空则把sqlserver中的cardID给strCardID
                        if (SqlServerdr != null)
                        {
                            strCardID = SqlServerdr["kh"].ToString();
                        }
                        DataRow Tempdr = DT1.Rows.Find(dr["EMP_CD"].ToString());
                        //如果SqlServerdr不为空说明sqlserver中存在此员工，继续判断是要Update还是Insert
                        //如果Tempdr不为空说明Mysql中存在此员工，继续判断是否需要Update
                        if (dr["EMP_STATUS"].ToString() == string.Empty)
                        { Istatus = 0; }
                        else
                        {
                            switch (Convert.ToInt32(dr["EMP_STATUS"].ToString()))
                            {
                                case 1:
                                case 8:
                                    Istatus = 0;
                                    break;
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                case 6:
                                case 7:
                                    Istatus = 1;
                                    break;
                                default:
                                    Istatus = 0;
                                    break;

                            }
                        }
                        if (Tempdr != null)
                        {
                            if (Tempdr["ID"].ToString() != dr["EMP_CD"].ToString() || Tempdr["Name"].ToString() != dr["EMP_NAME"].ToString() ||
                                Tempdr["depid"].ToString() != dr["DEPT_CD"].ToString() || Tempdr["depname"].ToString() != dr["DEPT_NAME"].ToString() ||
                                Tempdr["Direct"].ToString() != dr["D_I_CD"].ToString() || Tempdr["cardid"].ToString() != strCardID ||
                                Tempdr["costID"].ToString() != dr["EXP_DEPT"].ToString() || Tempdr["isOnWork"].ToString() != Istatus.ToString())
                            {
                                //if (dr["EMP_STATUS"].ToString() == string.Empty)
                                //{ Istatus = 0; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 0)
                                //{ Istatus = 1; }
                                //else if (Convert.ToInt32(dr["EMP_STATUS"].ToString()) == 1)
                                //{ Istatus = 0; }
                                strSQLUpdate = "update swipecard.testemployee SET Name = '" + dr["EMP_NAME"].ToString() +
                                                                              "',depid = '" + dr["DEPT_CD"].ToString() +
                                                                              "',depname = '" + dr["DEPT_NAME"].ToString() +
                                                                              "',Direct = '" + dr["D_I_CD"].ToString() +
                                                                              "',cardid = '" + strCardID +
                                                                              "',costID = '" + dr["EXP_DEPT"].ToString() +
                                                                              "',isOnWork = " + Istatus + "  ,updateDate =curdate() WHERE ID = '" + dr["EMP_CD"].ToString() + "' ";
                                cmd.CommandText = strSQLUpdate;
                                cmd.ExecuteNonQuery();
                                UpdateSumk++;
                                //WriteLog("-->Update112_EmpOK:" + UpdateSumk + strSQLUpdate + '\r', 1);
                            }
                        }
                        else
                        {
                            //如果Tempder为空，说明Mysql中要Insert
                            ++InsertSumk;
                            strSQLInsert += " ('" + dr["EMP_CD"].ToString() + "','"
                                                  + dr["EMP_NAME"].ToString() + "','"
                                                  + dr["DEPT_CD"].ToString() + "','"
                                                  + dr["DEPT_NAME"].ToString() + "','"
                                                  + dr["D_I_CD"].ToString() + "','"
                                                  + strCardID + "','"
                                                  + dr["EXP_DEPT"].ToString() + "',"
                                                  + "curdate()" + "),";
                            //WriteLog("-->Insert112_EmpOK:" + InsertSumk + " ('" + dr["EMP_CD"].ToString() + "','"
                            //                      + dr["EMP_NAME"].ToString() + "','"
                            //                      + dr["DEPT_CD"].ToString() + "','"
                            //                      + dr["DEPT_NAME"].ToString() + "','"
                            //                      + dr["D_I_CD"].ToString() + "','"
                            //                      + SqlServerdr["kh"].ToString() + "','"
                            //                      + dr["EXP_DEPT"].ToString() + "',"
                            //                      + "curdate()" + '\r', 1);

                            if (InsertSumk > 0 && InsertSumk % 2000 == 0)
                            {
                                strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                                cmd.CommandText = strSQLInsert;
                                cmd.ExecuteNonQuery();

                                strSQLInsert = "insert into swipecard.testemployee (ID,Name,depid,depname,Direct,cardid,costID,updateDate)  VALUES ";
                            }
                        }
                        //}
                    }
                    if (InsertSumk % 2000 != 0)
                    {
                        strSQLInsert = strSQLInsert.Substring(0, strSQLInsert.Length - 1);
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        tx.Commit();
                    }
                    else
                    {
                        tx.Commit();
                    }

                    WriteLog("-->Update112_EmployeeOK:" + UpdateSumk + '\r', 1);
                    WriteLog("-->insert112_EmployeeOK:" + InsertSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
                #endregion
            }
            

        }


        private void InsertMM(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            if (strBG == "CSBG")
            {
                string strSqlMySqlEmp = "SELECT CardID, SwipeCardTime FROM swipecard.modify_data2";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        strSQLInsert = "update swipecard.modify_data set SwipeCardTime2='" + Convert.ToDateTime(dr["SwipeCardTime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' where CardID='" + dr["CardID"].ToString() + "'";
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string strSqlMySqlEmp = "SELECT CardID, SwipeCardTime FROM swipecard.modify_data2";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        strSQLInsert = "update swipecard.modify_data set SwipeCardTime2='" + Convert.ToDateTime(dr["SwipeCardTime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' where CardID='" + dr["CardID"].ToString() + "'";
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                    }
                    tx.Commit();
                    WriteLog("-->Update112_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
 
        }


        private void InsertKK(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            if (strBG == "CSBG")
            {
                string strSqlMySqlEmp = "select  RecordID, SwipeCardTime2 from swipecard.modify_data where SwipeCardTime2 is not null";
                DataTable DT1 = MySqlHelp.QueryDBC(strSqlMySqlEmp);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        strSQLInsert = "update swipecard.testswipecardtime set SwipeCardTime2='" + Convert.ToDateTime(dr["SwipeCardTime2"].ToString()).ToString("yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' where RecordID='" + dr["RecordID"].ToString() + "'";
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                string strSqlMySqlEmp = "select  RecordID, SwipeCardTime2 from swipecard.modify_data where SwipeCardTime2 is not null";
                DataTable DT1 = MySqlHelp.QueryDbAsbg(strSqlMySqlEmp);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                //cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;

                int UpdateSumk = 0;
                int InsertSumk = 0;
                string strSQLUpdate = "";
                string strSQLInsert = "";
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        strSQLInsert = "update swipecard.testswipecardtime set SwipeCardTime2='" + Convert.ToDateTime(dr["SwipeCardTime2"].ToString()).ToString("yyyy/MM/dd HH:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + "' where RecordID='" + dr["RecordID"].ToString() + "'";
                        cmd.CommandText = strSQLInsert;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                    }
                    tx.Commit();
                    WriteLog("-->Update112_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert112_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
 
        }

        private void haha(string strBG)
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
            MySqlConnection MySqlConn = new MySqlConnection(connsql);
            MySqlConn.Open();
            MySqlTransaction tx = MySqlConn.BeginTransaction();
            MySqlCommand cmd = new MySqlCommand();
            cmd.CommandTimeout = 100;
            cmd.Connection = MySqlConn;
            cmd.Transaction = tx;

//            int UpdateSumk = 0;
//            int InsertSumk = 0;
            string strSQLUpdate = "";
//            string strSQLInsert = "insert into swipecard.emp_class VALUES ";
//            //找出testemployee中的人員在emp_class沒有當天的資料，則插入當天的資料，默認班別為4
//            strSQLUpdate = @"SELECT a.ID
//  FROM swipecard.emp_class a
// WHERE emp_date = curdate()-1
//GROUP BY a.ID, a.emp_date
//HAVING count(*) > 1";
//            DataTable DTID = MySqlHelp.QueryDBC(strSQLUpdate);
//            foreach (DataRow dr in DTID.Rows)
//            {
//                strSQLUpdate = @"DELETE FROM swipecard.emp_class
//  WHERE     emp_date = curdate() and update_time=curdate()
//       AND class_no = '4'
//       AND ID='" + dr["id"].ToString() + "'";
//                cmd.CommandText = strSQLUpdate;
//                cmd.ExecuteNonQuery();
//            }
            strSQLUpdate = "select emp.id from swipecard.testemployee emp where emp.isonwork = 0 and emp.id not in (select class.id from swipecard.emp_class class where class.emp_date = curdate()-1)";
            DataTable DTID = MySqlHelp.QueryDbAsbg(strSQLUpdate);
            foreach (DataRow dr in DTID.Rows)
            {
                strSQLUpdate = "insert into swipecard.emp_class values ('" + dr["ID"].ToString() + "',curdate()-1,'4',curdate()-1)";
                cmd.CommandText = strSQLUpdate;
                cmd.ExecuteNonQuery();
                //Insert4++;
            }

            tx.Commit();
            WriteLog("-->Update111_EmpClassOK:" + '\r', 1);
        }

        private void woriwori(string strBG)
        {
            MySqlHelper MySqlHelp = new MySqlHelper();
            string sql1 = "SELECT ID, cardid  FROM swipecard.testemployee where cardid!=''";
            if (strBG == "CSBG")
            {
                DataTable DT1 = MySqlHelp.QueryDBC(sql1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int UpdateSumk = 0;
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        string sql2 = "update swipecard.testswipecardtime set ID='" + dr["ID"].ToString() + "' where CardID='" + dr["cardid"].ToString() + "'";
                        cmd.CommandText = sql2;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                        WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
               

            }
            else if (strBG == "ASBG")
            {
                DataTable DT1 = MySqlHelp.QueryDbAsbg(sql1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int UpdateSumk = 0;
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        string sql2 = "update swipecard.testswipecardtime set ID='" + dr["ID"].ToString() + "' where CardID='" + dr["cardid"].ToString() + "'";
                        cmd.CommandText = sql2;
                        cmd.ExecuteNonQuery();
                        UpdateSumk++;
                        WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
 
        }

        private void Raw(string strBG)
        {
            MySqlHelper MySqlHelp = new MySqlHelper();
            string sql1 = "select recordid,id from swipecard.testswipecardtime where DATE_FORMAT(swipecardtime2, '%Y%m%d') = '20171018' and shift='N' and swipecardtime is null";
            if (strBG == "CSBG")
            {
                DataTable DT1 = MySqlHelp.QueryDBC(sql1);

                String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int UpdateSumk = 0;
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        string sql2 = "SELECT min(swipecardtime) time FROM swipecard.raw_record WHERE id = '" + dr["id"].ToString() + "' AND DATE_FORMAT(swipecardtime, '%Y%m%d %H:%M:%S') >= '20171017 19:00:00' AND  DATE_FORMAT(swipecardtime, '%Y%m%d %H:%M:%S') <= '20171017 20:01:00'";
                        DataTable DT2 = MySqlHelp.QueryDBC(sql2);
                        foreach (DataRow dr2 in DT2.Rows)
                        {
                            string sql3 = "UPDATE swipecard.testswipecardtime SET swipecardtime = '" + dr2["time"].ToString() + "' WHERE recordid = '" + dr["recordid"].ToString() + "'";
                            cmd.CommandText = sql3;
                            try
                            {
                                cmd.ExecuteNonQuery();
                            }
                            catch { continue; }
                            UpdateSumk++;
                        }
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
            else if (strBG == "ASBG")
            {
                DataTable DT1 = MySqlHelp.QueryDbAsbg(sql1);

                String connsql = "server=192.168.60.112;userid=root;password=foxlink;database=;charset=utf8";
                MySqlConnection MySqlConn = new MySqlConnection(connsql);
                MySqlConn.Open();
                MySqlTransaction tx = MySqlConn.BeginTransaction();
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandTimeout = 300;
                cmd.Connection = MySqlConn;
                cmd.Transaction = tx;
                int UpdateSumk = 0;
                try
                {
                    foreach (DataRow dr in DT1.Rows)
                    {
                        string sql2 = "SELECT min(swipecardtime) time FROM swipecard.raw_record WHERE id = '" + dr["id"].ToString() + "' AND DATE_FORMAT(swipecardtime, '%Y%m%d %H:%M:%S') >= '20171017 19:00:00' AND  DATE_FORMAT(swipecardtime, '%Y%m%d %H:%M:%S') <= '20171017 20:01:00'";
                        DataTable DT2 = MySqlHelp.QueryDbAsbg(sql2);
                        foreach (DataRow dr2 in DT2.Rows)
                        {
                            string sql3 = "UPDATE swipecard.testswipecardtime SET swipecardtime = '" + dr2["time"].ToString() + "' WHERE recordid = '" + dr["recordid"].ToString() + "'";
                            cmd.CommandText = sql3;
                            try
                            {
                                cmd.ExecuteNonQuery();
                            }
                            catch { continue; }
                            UpdateSumk++;
                        }
                    }
                    tx.Commit();
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
                }
                finally
                {
                    tx.Dispose();
                    MySqlConn.Close();
                }
            }
        }

        private void nima()
        {
            SqlHelp sqlHelp = new SqlHelp();
            MySqlHelper MySqlHelp = new MySqlHelper();
            ProgressHelp ProHelp = new ProgressHelp();

            String sql = "SELECT [編號] id, [姓名] name, [部門編號] depid, [部門名稱] depname, [直間接] direct, [卡號] cardid, [課部門編號] costid FROM V_RYBZ_TX WHERE [更新標誌]='N'"; // 查询语句
            //String sql = "select ygbh,fdate,bc from KQ_RECORD_BC where fdate>=CONVERT(varchar(100), GETDATE()-3, 111) and fdate<CONVERT(varchar(100), GETDATE()+2, 111) and ( " + strEMP + " )"; // 查询语句
            DataTable DT2 = sqlHelp.QuerySqlServerDB(sql);
            if (DT2 == null)
            {
                WriteLog("-->Query SqlServerdeptError" + '\r', 2);
                return;
            }
            else if (DT2.Rows.Count == 0)
            {
                WriteLog("-->Query SqlServerdeptSum:0" + '\r', 2);
                return;

            }

            String connsql = "server=192.168.60.111;userid=root;password=foxlink;database=;charset=utf8";
            //String connsql = "server=10.64.155.200;userid=root;password=mysql;database=mysql;charset=utf8";
            MySqlConnection MySqlConn = new MySqlConnection(connsql);
            MySqlConn.Open();
            MySqlTransaction tx = MySqlConn.BeginTransaction();
            MySqlCommand cmd = new MySqlCommand();
            //cmd.CommandTimeout = 300;
            cmd.Connection = MySqlConn;
            cmd.Transaction = tx;

            int UpdateSumk = 0;
            int InsertSumk = 0;
            int Istatus = 0;
            string strSQLUpdate = "";
            string strSQLUpdateIswork = "";

            foreach (DataRow dr in DT2.Rows)
            {
                string sql22 = "update swipecard.testemployee set isOnWork=0 where ID='" + dr["id"].ToString() + "'";
                cmd.CommandText = sql22;
                cmd.ExecuteNonQuery();
                UpdateSumk++;

            }
            tx.Commit();

            WriteLog("-->Update111_EmployeeOK:" + UpdateSumk + '\r', 1);
 
        }

        private void meimei()
        {
            MySqlHelper MySqlHelp = new MySqlHelper();
            string sql1 = @"SELECT ID,
       SwipeCardTime,
       SwipeCardTime2,
       CheckState,
       PROD_LINE_CODE,
       WorkshopNo,
       PRIMARY_ITEM_NO,
       RC_NO,
       Shift
  FROM swipecard.testswipecardtime
 WHERE     (SwipeCardTime >= '2018/07/05' OR SwipeCardTime2 >= '2018/07/05') AND CheckState = 0 and id IN ('1456958','132716') ";

            DataTable DT1 = MySqlHelp.QueryDBC(sql1);

            string connsql = "Data Source=192.168.64.224/rtodb;User ID=swipe;Password=mis_realt";
            //string connsql = "Data Source=192.168.144.187/scard;User ID=swipe;Password=mis_swipe";
            OracleConnection OracleConn = new OracleConnection(connsql);
            OracleConn.Open();
            OracleTransaction tx = OracleConn.BeginTransaction(); ;
            OracleCommand cmd = new OracleCommand();
            //cmd.CommandTimeout = 60;
            cmd.Connection = OracleConn;
            cmd.Transaction = tx;
            int UpdateSumk = 0;
            int InsertSumk = 0;
            string strSQLUpdate = "";
            string strSQLInsert = "insert all into swipe.emp_class VALUES ";
            bool b = true;

            try
            {
                foreach (DataRow dr in DT1.Rows)
                {
                    string strSwipeData = "";

                   string sa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr["SwipeCardTime"].ToString()).ToString("yyyy-MM-dd") + "'";
                   string sb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr["SwipeCardTime2"].ToString()).ToString("yyyy-MM-dd") + "'";
                   if (sa == "NULL" && sb != "NULL") 
                   {
                       if (dr["Shift"].ToString() == "D")
                       {
                           strSwipeData = sb;
                       }
                       else if (dr["Shift"].ToString() == "N")
                       {
                           strSwipeData = "'" + Convert.ToDateTime(dr["SwipeCardTime2"].ToString()).AddDays(-1).ToString("yyyy-MM-dd") + "'";
                       }
                   }
                   else if (sa != "NULL") 
                   {
                      
                       strSwipeData = sa;

                   }
                   string sql = "select class_no from swipecard.emp_class where emp_date=" + strSwipeData + " and ID='" + dr["ID"].ToString() + "'";
                   string strClassNO = MySqlHelp.QueryDBC(sql).Rows[0][0].ToString();



                  

                   string saa = dr[1].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr["SwipeCardTime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                   string sbb = dr[2].Equals(DBNull.Value) ? "NULL" : "'" + Convert.ToDateTime(dr["SwipeCardTime2"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'";


                   string strSQL = @"INSERT INTO SWIPE.CSR_SWIPECARDTIME  VALUES  ('" + dr["ID"].ToString() + "',"
                                                                                 + strSwipeData + ",to_date("
                                                                                 + saa + ",'yyyy-mm-dd hh24:mi:ss') ,to_date("
                                                                                 + sbb + ",'yyyy-mm-dd hh24:mi:ss'),'"
                                                                                 + dr["CheckState"].ToString() + "','"
                                                                                 + dr["PROD_LINE_CODE"].ToString() + "','"
                                                                                 + dr["WorkshopNo"].ToString() + "','"
                                                                                 + dr["PRIMARY_ITEM_NO"].ToString() + "','"
                                                                                 + dr["RC_NO"].ToString() + "','"
                                                                                 + dr["Shift"].ToString() + "','"
                                                                                 + strClassNO + "')";
                   cmd.CommandText = strSQL;
                    cmd.ExecuteNonQuery();
                    UpdateSumk++;
                    WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
                }
                tx.Commit();
                WriteLog("-->Update111_EmpClassOK:" + UpdateSumk + '\r', 1);
            }
            catch (Exception ex)
            {
                tx.Rollback();
                WriteLog("-->insert111_EmpClassError,Rollback :" + ex.Message + '\r', 2);
            }
            finally
            {
                tx.Dispose();
                OracleConn.Close();
            }

 
        }

    }
    
}
