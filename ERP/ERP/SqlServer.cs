using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Combination
{
    class SqlServer
    {
        SqlConnection conn = new SqlConnection();
        string DbName = "ChengyiYuntech";
        DataSet ds;
        public void Connect(string DbName)
        {
            this.DbName = DbName;
            string path = @"Data Source=192.168.1.252;Initial Catalog=ChengyiYuntech;User ID=SA;Password=chengyi";
            if (conn.State.ToString() == "Open")
                conn.Close();
            conn.ConnectionString = path;
            conn.Open();
        }


        public void Close()
        {
            conn.Close();
        }

        public DataSet SqlCmd(string tableCmd)
        {
            Connect(DbName);
            SqlDataAdapter da = new SqlDataAdapter(tableCmd, conn);
            ds = new DataSet();
            ds.Clear();
            da.Fill(ds);
            return ds;
        }
    }
}
