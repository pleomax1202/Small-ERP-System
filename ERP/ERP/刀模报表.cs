using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Combination
{
    public partial class 刀模报表 : Form
    {
        public 刀模报表()
        {
            InitializeComponent();
            txtKCode.Focus();
        }

        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;
        Sql sql = new Sql();

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (bunifuCheckbox1.Checked == true)
            {
                if(txtCode.Text == "")
                {
                    dataGridView2.Rows.Clear();
                    dataGridView2.Columns[3].Visible = false;
                    dataGridView2.Columns[0].HeaderText = "印刷版号";
                    dataGridView2.Columns[1].HeaderText = "数量";
                    dataGridView2.Columns[2].HeaderText = "单位";

                    string query = @"select pNo as '印刷版号',sum(ppqty) as '数量' ,'张' AS '单位' from (
                                    (SELECT O_001 AS PNO,b.oid,omachinecode,opname,odate,
                                    pppcs/F_110  as ppqty
                                    FROM [ChengyiYuntech].[dbo].[ScanRecord]a,
                                    [ChengyiYuntech].[dbo].[ProduceOrder]b,
                                    [ChengyiYuntech].[dbo].[Machine]c,
                                    ["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_002 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno ) " +
                                    "UNION ALL " +
                                    "(SELECT O_003 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno ) " +
                                    "UNION ALL " +
                                    "(SELECT O_004 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno ) " +
                                    "UNION ALL " +
                                    "(SELECT O_005 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno ) " +
                                    "UNION ALL " +
                                    "(SELECT O_006 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno ) " +
                                    "UNION ALL " +
                                    "(SELECT O_007 AS PNO,b.oid,omachinecode,opname,odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM [ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode,2) ='BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    ") A  " +
                                    "WHERE  PNO <>'' AND PNO IS NOT NULL " +
                                    "group by pNo";

                    Load_Data(query);
                }
                else if(txtCode.Text != "")
                {
                    dataGridView2.Rows.Clear();
                    dataGridView2.Columns[3].Visible = true;
                    dataGridView2.Columns[0].HeaderText = "日期";
                    dataGridView2.Columns[1].HeaderText = "产品名称";
                    dataGridView2.Columns[2].HeaderText = "数量";

                    string query = @"select * from 
                                    (select case when grouping(odate) = 1 and grouping(opname) = 0 then '小计' else convert(varchar(10), odate, 120) end as odate, OpName, sum(ppqty) as ppqty, '张' AS unit from(
                                    (SELECT O_001 AS PNO, b.oid, omachinecode, opname, odate,
                                    pppcs/F_110  as ppqty
                                    FROM[ChengyiYuntech].[dbo].[ScanRecord]a,
                                    [ChengyiYuntech].[dbo].[ProduceOrder]b,
                                    [ChengyiYuntech].[dbo].[Machine]c,
                                    ["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_002 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_003 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_004 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_005 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_006 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    "UNION ALL " +
                                    "(SELECT O_007 AS PNO, b.oid, omachinecode, opname, odate, " +
                                    "pppcs/F_110  as ppqty " +
                                    "FROM[ChengyiYuntech].[dbo].[ScanRecord]a, " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder]b, " +
                                    "[ChengyiYuntech].[dbo].[Machine]c, " +
                                    "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                    "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                    "where left(omachinecode, 2) = 'BB' " +
                                    "and a.poid = b.id " +
                                    "and omachinecode = c.Mcode " +
                                    "and d.FitemID = e.FitemID " +
                                    "and b.oid = d.fbillno) " +
                                    ") A " +
                                    "WHERE  PNO <> '' AND PNO IS NOT NULL " +
                                    "and pno = '"+ txtCode.Text +"' " +
                                    "group by pNo,OPNAME,odate with rollup) as finalTable where odate is not null and ppqty<> 0";

                    Load_Data(query);
                }
            }
            else
            {
                if (txtKCode.Text == "")
                {
                    dataGridView1.Columns[0].HeaderText = "刀模编号";
                    dataGridView1.Columns[1].HeaderText = "产品名称";
                    dataGridView1.Columns[2].Visible = false;
                    dataGridView1.Columns[3].Visible = false;

                    string query = @"select a.OCcode,c.Fmodel,SUM(b.PPQty) as PPQty,SUM(b.PPPcs) as PCSQty from [ChengyiYuntech].[dbo].[ProduceOrder] a
                                ,[ChengyiYuntech].[dbo].[ScanRecord] b
                                ,[" + sql.CYDB + "].[dbo].[T_Icitem] c " +
                                    ",[" + sql.CYDB + "].[dbo].[ICMO] d " +
                                    ",[ChengyiYuntech].[dbo].[Machine] e " +
                                    "where a.ID = b.POID and a.OStatus = '1' " +
                                    "and a.OID = d.Fbillno " +
                                    "and d.FitemID = c.FitemID " +
                                    "and a.OMachineCode = e.MCode  " +
                                    "and e.Mcode like 'C%' " +
                                    "and a.OCcode <> '' " +
                                    "and OCcode <> '07.K.01.005' " +
                                    "group by a.OCcode,c.Fmodel Order by a.OCcode";

                    Load_Data(query);
                }
                else if (txtKCode.Text != "")
                {
                    dataGridView1.Columns[0].HeaderText = "日期";
                    dataGridView1.Columns[1].HeaderText = "机台编号";
                    dataGridView1.Columns[2].Visible = true;
                    dataGridView1.Columns[3].Visible = true;

                    string query = @"(select CONVERT(char(10), a.Odate, 23) as ODate,c.Mcode,c.Mname,a.OCCode,SUM(b.PPQty) as N'數量(張)',SUM(b.PPPcs) as N'數量(PCS)' from 
                                [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c
                                where a.ID = b.POID and a.OCCode <> '' and c.Mcode =a.OMachineCode and a.OCCode = '" + txtKCode.Text + "' group by a.OCCode ,CONVERT(char(10), a.Odate, 23),c.Mcode,c.Mname) union " +
                                    "(select '3000-12-17' as ODate,'小计' as Mcode,'' as Mname,'' as OCCode,SUM(b.PPQty) as N'數量(張)',SUM(b.PPPcs) as N'數量(PCS)' from  " +
                                    "[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c " +
                                    "where a.ID = b.POID and a.OCCode <> '' and c.Mcode =a.OMachineCode and a.OCCode = '" + txtKCode.Text + "' group by a.OCCode)";

                    Load_TitleData();
                    Load_Data(query);

                    int count = dataGridView1.Rows.Count;
                    dataGridView1.Rows[count - 1].Cells[0].Value = "";
                }
            }
            Cursor = Cursors.Default;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

                Export_Data();


        }

        private void Load_Data(string query)
        {
            dataGridView1.Rows.Clear();

            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            if(bunifuCheckbox1.Checked == true)
            {
                if(txtCode.Text == "")
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dataGridView2.Rows.Add();
                        dataGridView2.Rows[n].Cells[0].Value = item["印刷版号"].ToString();
                        dataGridView2.Rows[n].Cells[1].Value = Convert.ToDecimal(item["数量"]).ToString("N0");
                        dataGridView2.Rows[n].Cells[2].Value = item["单位"].ToString();
                    }
                }
                else if(txtCode.Text != "")
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dataGridView2.Rows.Add();
                        dataGridView2.Rows[n].Cells[0].Value = item["odate"].ToString();
                        dataGridView2.Rows[n].Cells[1].Value = item["OpName"].ToString();
                        dataGridView2.Rows[n].Cells[2].Value = Convert.ToDecimal(item["ppqty"]).ToString("N0");
                        dataGridView2.Rows[n].Cells[3].Value = item["unit"].ToString();
                    }
                }
            }
            else
            {
                if (txtKCode.Text == "")
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["OCcode"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["Fmodel"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["PPQty"]).ToString("N0");
                        dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["PCSQty"]).ToString("N0");
                    }
                }
                if (txtKCode.Text != "")
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["ODate"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["Mcode"].ToString();
                        dataGridView1.Rows[n].Cells[2].Value = item["Mname"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["OCCode"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["數量(張)"]).ToString("N0");
                        dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["數量(PCS)"]).ToString("N0");
                    }
                }
            }
        }

        private void Export_Data()
        {
            if(bunifuCheckbox1.Checked == true)
            {
                if (txtCode.Text == "")
                {
                    exportPath = path + @"\版号汇总表导出";
                    filePath = inputPath + @"\版号汇总表";
                }
                else if (txtCode.Text != "")
                {
                    exportPath = path + @"\版号明细表导出";
                    filePath = inputPath + @"\版号明细表";
                }
            }
            else
            {
                if (txtKCode.Text != "")
                {
                    exportPath = path + @"\刀模明细表导出";
                    filePath = inputPath + @"\刀模明细表";
                }
                else if (txtKCode.Text == "")
                {
                    exportPath = path + @"\刀模汇总表导出";
                    filePath = inputPath + @"\刀模汇总表";
                }
            }

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();

            // 讓Excel文件可見
            excelApp.Visible = false;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);

            wBook = excelApp.Workbooks.Open(filePath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

            // 設定活頁簿焦點
            wBook.Activate();

            wSheet = (Excel._Worksheet)wBook.Worksheets[1];


            if(bunifuCheckbox1.Checked == true)
            {
                if (txtCode.Text == "")
                {
                    wSheet.Name = "版号汇总表";

                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 4, 1] = Convert.ToString(dataGridView2.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 4, 2] = Convert.ToString(dataGridView2.Rows[i].Cells[1].Value);
                        wSheet.Cells[i + 4, 3] = Convert.ToString(dataGridView2.Rows[i].Cells[2].Value);
                    }
                }
                else if (txtCode.Text != "")
                {
                    wSheet.Name = "版号明細表";
                    wSheet.Cells[3, 2] = txtCode.Text;

                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 5, 1] = Convert.ToString(dataGridView2.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 5, 2] = Convert.ToString(dataGridView2.Rows[i].Cells[1].Value);
                        wSheet.Cells[i + 5, 3] = Convert.ToString(dataGridView2.Rows[i].Cells[2].Value);
                        wSheet.Cells[i + 5, 4] = Convert.ToString(dataGridView2.Rows[i].Cells[3].Value);
                    }
                }
            }
            else
            {
                if (txtKCode.Text != "")
                {
                    wSheet.Name = "刀模明细表";
                    wSheet.Cells[3, 2] = txtKCode.Text;
                    wSheet.Cells[3, 4] = txtKName.Text;

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        {
                            wSheet.Cells[i + 5, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                        }
                    }
                }
                else if (txtKCode.Text == "")
                {
                    wSheet.Name = "刀模汇总表";

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 4, 1] = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 4, 2] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                        wSheet.Cells[i + 4, 3] = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                        wSheet.Cells[i + 4, 4] = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                    }
                }
            }
            
            

            Excel.Range last = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range allRange = wSheet.get_Range("A1", last);
            allRange.Font.Size = "14";
            allRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;        //格線
            allRange.Columns.AutoFit();

            //Save Excel
            wBook.SaveAs(exportPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;

            GC.Collect();
            MessageBox.Show("导出成功");
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            Load_TitleData();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = btnSeek;
            }
        }

        private void Load_TitleData()
        {
            if (txtKCode.Text != "")
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(@"select Fmodel from ["+ sql.CYDB +"].[dbo].[T_Icitem] where F_111 = '" + txtKCode.Text + "' or F_125 = '" + txtKCode.Text + "' and Fmodel not like '%刀模%'");

                foreach (DataRow item in dt.Rows)
                {
                    txtKName.Text = item["Fmodel"].ToString();
                }
            }
        }

        private void bunifuCheckbox1_OnChange(object sender, EventArgs e)
        {
            if(bunifuCheckbox1.Checked == true)
            {
                label1.Text = "柔印版号";
                lblKName.Visible = false;
                txtKName.Visible = false;
                txtKCode.Visible = false;
                txtCode.Visible = true;
                dataGridView1.Visible = false;
                dataGridView2.Visible = true;
            }
            else
            {
                label1.Text = "刀模版号";
                lblKName.Visible = true;
                txtKName.Visible = true;
                txtKCode.Visible = true;
                txtCode.Visible = false;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
            }
        }
    }
}
