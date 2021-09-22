using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Combination
{
    class Sql
    {
        DataTable dt;
        DataSet ds;
        SqlConnection con = new SqlConnection(@"Data Source=localhost;Initial Catalog=ChengyiYuntech;Integrated Security=True");

        public string CYDB = "AIS20190520180820";
        public string CKDB = "AIS20190520181425";

        public DataTable getQuery(string Query)
        {
            SqlDataAdapter sda = new SqlDataAdapter(Query, con);

            dt = new DataTable();
            dt.Clear();
            sda.Fill(dt);
            return dt;
        }

        public DataSet SqlCmdDS(string tableCmd)
        {
            SqlDataAdapter da = new SqlDataAdapter(tableCmd, con);
            ds = new DataSet();
            ds.Clear();
            da.Fill(ds);
            return ds;
        }

        public void sqlCmd(string query)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
