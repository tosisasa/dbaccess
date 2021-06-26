using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.Common;
using System.Data.OleDb;

namespace XXX
{
    public class DBAccess
    {
        //DB接続関連はクラスで保有
        DbProviderFactory factory;
        DbConnection conn = null;
        DbConnectionStringBuilder ocsb;
        DbCommand cmd;
        DbDataAdapter da;
        System.Data.Common.DbTransaction trans;


        public String sExceptMsg;
        public String sExceptStackTrace;


        public bool connectDB(String schema)
        {
            sExceptMsg = null;
            sExceptStackTrace = null;

            bool isConnectDB = true;

            try
            {
                bool bOracle = false;
                switch (schema)
                {
                    case ("ユーザ名@TNS名"):
                        //Oracle
                        bOracle = true;
                        break;
                    default:
                        break;
                }
                if (bOracle && Param.getValue("DB_PROVIDER_ORACLE").Equals("ODP.NET"))
                {
                    factory = DbProviderFactories.GetFactory("Oracle.DataAccess.Client");
                    ocsb = factory.CreateConnectionStringBuilder();

                }
                else
                {
                    factory = DbProviderFactories.GetFactory("System.Data.OleDb");
                    ocsb = factory.CreateConnectionStringBuilder();

                    if (bOracle)
                    {
                        //Oracle
                        ocsb["Provider"] = "OraOLEDB.Oracle.1";
                        ocsb["OLE DB Services"] = "-2"; //接続プール無効

                    }
                    else
                    {
                        //ACCDB(Access)
                        ocsb["Provider"] = "Microsoft.ACE.OLEDB.12.0";
                        ocsb["Mode"] = "Read"; //読み取り専用
                    }
                }


                //接続文字列の設定
                if (schema == "accdbファイル名")
                {
                    if (!Param.getValue("DBTEST").Equals("1"))
                    {
                        //本番サーバへ接続する。
                        ocsb["User ID"] = "ユーザ名";
                        ocsb["password"] = "パスワード";
                        ocsb["Data Source"] = "TNS名";
                    }
                    else
                    {
                        //テストサーバへ接続する。

                    }
                }
                else
                {
                    //error
                }

                conn = factory.CreateConnection();
                conn.ConnectionString = ocsb.ConnectionString;


                //データベース接続
                conn.Open();
                trans = conn.BeginTransaction();


                cmd = conn.CreateCommand();
                cmd.Transaction = trans;
            }
            catch (Exception e)
            {
                sExceptMsg = e.Message;
                sExceptStackTrace = e.StackTrace;

                isConnectDB = false;
            }

            return isConnectDB;

        }

        public void closeDB(bool isCommit)
        {
            if (isCommit)
            {
                try
                {
                    trans.Commit();
                }
                catch (Exception e)
                {

                }
            }
            else
            {
                try
                {
                    trans.Rollback();
                }
                catch (Exception e)
                {

                }
            }
            try
            {
                conn.Close();
            }
            catch (Exception e)
            {

            }
        }

        public DataSet selectDB(String sql, String schema)
        {
            DataSet ds = null;
            connectDB(schema);
            try
            {
                cmd.CommandText = sql;

                da = factory.CreateDataAdapter();
                da.SelectCommand = cmd;

                //SELECTの実行
                ds = new DataSet();
                da.Fill(ds);
            }
            finally
            {
                closeDB(true);
            }

            return ds;
        }


        public DataSet selectDBOnTrans(String sql)
        {
            DataSet ds = null;
            try
            {
                cmd.CommandText = sql;

                da = factory.CreateDataAdapter();
                da.SelectCommand = cmd;

                //SELECTの実行
                ds = new DataSet();
                da.Fill(ds);
            }
            finally
            {
            }

            return ds;
        }


        public bool updateDB(String sql, String schema)
        {
            bool bRet = false;

            connectDB(schema);
            try
            {
                cmd.CommandText = sql;

                //UPDATE実行
                cmd.ExecuteNonQuery();
                bRet = true;
            }
            finally
            {
                closeDB(true);
            }
            return bRet;


        }

        /**
         * message=falseのときは、例外を抑止する。
         * 
         **/
        public bool updateDBOnTrans(String sql, bool message)
        {

            try
            {
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception e)
            {
                if (message)
                {
                    throw e;
                }

            }
            return false;

        }

    }
}


