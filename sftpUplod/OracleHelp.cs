using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using Oracle.ManagedDataAccess.Client;

namespace sftpUplod
{
    class OracleHelp
    {
        public DataTable QueryDbCSBG(string sqlStr)
        {
            try
            {
                string connectionString = "Data Source=192.168.64.224/rtodb;User ID=usyn;Password=mis_syn";
                //string connectionString = "Data Source=192.168.144.187/rtodb;User ID=swipe;Password=mis_swipe";
                using (OracleConnection sqlconn = new OracleConnection(connectionString))
                {
                    //sqlconn.Open();
                    OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "datatable1");
                    DataTable dt1 = ds1.Tables[0];
                    return dt1;
                }
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now + "-->Querydb111 Error:" + ex.Message);
                return null;
            }

        }

        public DataTable QueryDbASBG(string sqlStr)
        {
            string connectionString = "";
            try
            {
                using (OracleConnection sqlconn = new OracleConnection(connectionString))
                {
                    //sqlconn.Open();
                    OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "datatable1");
                    DataTable dt1 = ds1.Tables[0];
                    return dt1;
                }
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now + "-->Querydb112 Error:" + ex.Message);
                return null;
            }
        }

        public DataTable QueryDb65_230(string sqlStr)
        {
            string connectionString = "";
            try
            {
                using (OracleConnection sqlconn = new OracleConnection(connectionString))
                {
                    //sqlconn.Open();
                    OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "datatable1");
                    DataTable dt1 = ds1.Tables[0];
                    return dt1;
                }
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now + "-->Querydb65_230 Error:" + ex.Message);
                return null;
            }
        }



        //public DataTable QueryDB(string sqlStr, List<OracleParameter> args)
        //{
        //   // string connectionString = ConfigurationManager.ConnectionStrings["conDSGDB"].ConnectionString.ToString();
        //    using (OracleConnection sqlconn = new OracleConnection(connectionString))
        //    {
        //        //sqlconn.Open();
        //        OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
        //        if ((args != null) && (args.Count > 0))
        //        {
        //            da1.SelectCommand.Parameters.Clear();
        //            da1.SelectCommand.InitialLONGFetchSize = -1;
        //            //當參數順序不等於在SQL指令內出現的順序時,只要參數名稱對,就會自己找上
        //            //但也會造成原本是參數名稱對不上,但順序對上的那一種出錯
        //            //oda.SelectCommand.BindByName = true;
        //            foreach (OracleParameter p in args)
        //            {
        //                da1.SelectCommand.Parameters.Add(p);
        //            }
        //        }
        //        DataTable table = new DataTable("datatable1");
        //        da1.Fill(table);
        //        return table;
        //    }


        //}
        //public DataTable QueryDB(string sqlStr, List<OracleParameter> args, string connectionString)
        //{
        //    // string connectionString = ConfigurationManager.ConnectionStrings["conOracleDB"].ConnectionString.ToString();
        //    using (OracleConnection sqlconn = new OracleConnection(connectionString))
        //    {
        //        //sqlconn.Open();
        //        OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
        //        if ((args != null) && (args.Count > 0))
        //        {
        //            da1.SelectCommand.Parameters.Clear();
        //            da1.SelectCommand.InitialLONGFetchSize = -1;
        //            //當參數順序不等於在SQL指令內出現的順序時,只要參數名稱對,就會自己找上
        //            //但也會造成原本是參數名稱對不上,但順序對上的那一種出錯
        //            //oda.SelectCommand.BindByName = true;
        //            foreach (OracleParameter p in args)
        //            {
        //                da1.SelectCommand.Parameters.Add(p);
        //            }
        //        }
        //        DataTable table = new DataTable("datatable1");
        //        da1.Fill(table);
        //        return table;
        //    }
        //}
        //public DataTable QueryDB86(string sqlStr, List<OracleParameter> args)
        //{
        //   // string connectionString = ConfigurationManager.ConnectionStrings["conProductionDB"].ConnectionString.ToString();
        //    using (OracleConnection sqlconn = new OracleConnection(connectionString))
        //    {
        //        //sqlconn.Open();
        //        OracleDataAdapter da1 = new OracleDataAdapter(sqlStr, sqlconn);
        //        if ((args != null) && (args.Count > 0))
        //        {
        //            da1.SelectCommand.Parameters.Clear();
        //            da1.SelectCommand.InitialLONGFetchSize = -1;
        //            //當參數順序不等於在SQL指令內出現的順序時,只要參數名稱對,就會自己找上
        //            //但也會造成原本是參數名稱對不上,但順序對上的那一種出錯
        //            //oda.SelectCommand.BindByName = true;
        //            foreach (OracleParameter p in args)
        //            {
        //                da1.SelectCommand.Parameters.Add(p);
        //            }
        //        }
        //        DataTable table = new DataTable("datatable1");
        //        da1.Fill(table);
        //        return table;
        //    }


        //}

        //public int executeUpdate(string strSQL, OracleConnection pconn, OracleTransaction trans, List<OracleParameter> args)
        //{
        //    int i = 0;
        //    using (OracleCommand cmd = new OracleCommand(strSQL, pconn))
        //    {
        //        foreach (OracleParameter para in args)
        //        {
        //            cmd.Parameters.Add(para);
        //        }
        //        cmd.Transaction = trans;
        //        i = cmd.ExecuteNonQuery();
        //    }
        //    return i;
        //}

        private void WriteLog(String strLog)
        {
            DateTime currenttime = System.DateTime.Now;
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
                string LogStr = currenttime + " " + strLog;
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
    }
}
