using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;

namespace DataTable2Excel
{
    public class SQLiteHellper
    {
        public static string ConnectionString = "Data Source= " + AppDomain.CurrentDomain.BaseDirectory + "DataBase\\Data.db" + ";Pooling=true;FailIfMissing=false";
        public static DataSet ExecuteQuery(string strSql, params object[] p)
        {
            using (SQLiteConnection conn = new SQLiteConnection(ConnectionString))
            {
                using (SQLiteCommand command = new SQLiteCommand(strSql,conn))
                {

                    DataSet ds = new DataSet();
                    try
                    {
                        //PrepareCommand(command, conn, cmdText, p);
                        SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                        da.Fill(ds);
                        return ds;
                    }
                    catch (Exception ex)
                    {
                        return ds;
                    }
                }
            }
        }


    }
}
