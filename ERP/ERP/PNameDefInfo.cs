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
    public partial class PNameDefInfo : Form
    {
        Sql sql = new Sql();
        string dropDown;
        string query;
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;

        public PNameDefInfo()
        {
            InitializeComponent();
        }


        private void PNameDefInfo_Load(object sender, EventArgs e)
        {
            ddValue.SelectedIndex = 0;
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Load_Data();
            Cursor = Cursors.Default;
        }


        private void btnExport_Click(object sender, EventArgs e)
        {
            Export_Data();
        }

        private void Export_Data()
        {
            if(ddValue.SelectedIndex == 0)
            {
                exportPath = path + @"\成品自定义报表导出";
                filePath = inputPath + @"\成品自定义报表";
            }
            else if(ddValue.SelectedIndex == 1 || ddValue.SelectedIndex == 2)
            {
                exportPath = path + @"\刀模柔印版自定义报表导出";
                filePath = inputPath + @"\刀模柔印版自定义报表";
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

            if (ddValue.SelectedIndex == 0)
            {
                wSheet.Name = "成品自定义报表";
                for (int i = 0; i < dgvPNameDef.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvPNameDef.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dgvPNameDef.Rows[i].Cells[j].Value);
                    }
                }
            }
            else if (ddValue.SelectedIndex == 1)
            {
                wSheet.Name = "刀模柔印版自定义";
                for (int i = 0; i < dgv17A.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv17A.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dgv17A.Rows[i].Cells[j].Value);
                    }
                }
            }
            else if (ddValue.SelectedIndex == 2)
            {
                wSheet.Name = "刀模柔印版自定义";
                for (int i = 0; i < dgv17F.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv17F.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dgv17F.Rows[i].Cells[j].Value);
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

        private void Load_Data()
        {
            DataTable dt = new DataTable();

            if (ddValue.SelectedIndex == 0)
            {
                dgvPNameDef.Rows.Clear();
                dgvPNameDef.Visible = true;
                dgv17A.Visible = false;
                dgv17F.Visible = false;

                dropDown = "14.A";

                query = @"select * from ( 
                            (select a.FNumber, a.FModel, SUM(b.FQty / c.Fcoefficient) as FQty, a.FNetWeight, a.FGrossWeight, a.F_102, a.F_108, a.F_109, a.F_110, a.F_146, a.F_123, a.F_114, a.F_129
                            from[" + sql.CYDB + "].[dbo].[T_ICItem] a,[" + sql.CYDB + "].[dbo].[ICInventory] b,[" + sql.CYDB + "].[dbo].[t_measureunit] c " +
                            "where a.FitemID = b.FitemID and a.Fdeleted = '0' and a.FDefaultLoc = b.FstockID and(b.FstockID = '809' or b.FstockID = '810' or b.FstockID = '20421') " +
                            "and(a.FstoreUnitID = c.FMeasureUnitID and a.FUnitGroupID = c.FunitGroupID) " +
                            "group by a.FNumber, a.FModel, a.FNetWeight, a.FGrossWeight, a.F_102, a.F_108, a.F_109, a.F_110, a.F_146, a.F_123, a.F_114, a.F_129)union  " +
                            "(select a.FNumber, a.FModel, SUM(b.FQty / c.Fcoefficient) as FQty, a.FNetWeight, a.FGrossWeight, a.F_102, a.F_108, a.F_109, a.F_110, a.F_146, a.F_123, a.F_114, a.F_129  " +
                            "from[" + sql.CKDB + "].[dbo].[T_ICItem] a,[" + sql.CKDB + "].[dbo].[ICInventory] b,[" + sql.CKDB + "].[dbo].[t_measureunit] c " +
                            "where a.FitemID = b.FitemID and a.Fdeleted = '0' and a.FDefaultLoc = b.FstockID and b.FstockID = '19762' " +
                            "and(a.FstoreUnitID = c.FMeasureUnitID and a.FUnitGroupID = c.FunitGroupID) " +
                            "group by a.FNumber, a.FModel, a.FNetWeight, a.FGrossWeight, a.F_102, a.F_108, a.F_109, a.F_110, a.F_146, a.F_123, a.F_114, a.F_129))a " +
                            "where a.FNumber like '%" + txtValue.Text + "%' and a.FNumber like '" + dropDown + "%'";

                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dgvPNameDef.Rows.Add();
                    dgvPNameDef.Rows[n].Cells[0].Value = item["FNumber"].ToString();
                    dgvPNameDef.Rows[n].Cells[1].Value = item["FModel"].ToString();
                    dgvPNameDef.Rows[n].Cells[2].Value = Convert.ToDecimal(item["FQty"]).ToString("N0");
                    dgvPNameDef.Rows[n].Cells[3].Value = Convert.ToDecimal(item["FNetWeight"]).ToString("N5");
                    dgvPNameDef.Rows[n].Cells[4].Value = Convert.ToDecimal(item["FGrossWeight"]).ToString("N5");
                    dgvPNameDef.Rows[n].Cells[5].Value = item["F_102"].ToString();
                    dgvPNameDef.Rows[n].Cells[6].Value = item["F_108"].ToString();
                    dgvPNameDef.Rows[n].Cells[7].Value = item["F_109"].ToString();
                    dgvPNameDef.Rows[n].Cells[8].Value = item["F_110"].ToString();
                    dgvPNameDef.Rows[n].Cells[9].Value = item["F_146"].ToString();
                    dgvPNameDef.Rows[n].Cells[10].Value = Convert.ToDecimal(item["F_123"]).ToString("N5");
                    dgvPNameDef.Rows[n].Cells[11].Value = Convert.ToDecimal(item["F_114"]).ToString("N5");
                    dgvPNameDef.Rows[n].Cells[12].Value = Convert.ToDecimal(item["F_129"]).ToString("N5");
                }
            }
            else if(ddValue.SelectedIndex == 1)
            {
                dgv17A.Rows.Clear();
                dgvPNameDef.Visible = false;
                dgv17A.Visible = true;
                dgv17F.Visible = false;

                dropDown = "17.A";

                query = @"select * from(
                        (select FModel,FNumber,isnull(F_181,'') as 斷張刀,isnull(F_111,'') as 編號,'刀模1' as 自定義位置 from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_181,'') as 斷張刀,isnull(F_125,'') as 編號,'刀模2' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') " +
                        ")a where a.Fnumber like '%" + txtValue.Text + "%' order by a.Fmodel";

                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dgv17A.Rows.Add();
                    dgv17A.Rows[n].Cells[0].Value = item["FModel"].ToString();
                    dgv17A.Rows[n].Cells[1].Value = item["FNumber"].ToString();
                    dgv17A.Rows[n].Cells[2].Value = item["斷張刀"].ToString();
                    dgv17A.Rows[n].Cells[3].Value = item["編號"].ToString();
                    dgv17A.Rows[n].Cells[4].Value = item["自定義位置"].ToString();
                }
            }
            else if(ddValue.SelectedIndex == 2)
            {
                dgv17F.Rows.Clear();
                dgvPNameDef.Visible = false;
                dgv17A.Visible = false;
                dgv17F.Visible = true;

                dropDown = "17.F";

                query = @"select * from (
                        (select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_182,'') as 編號,'柔印版1' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_183,'') as 編號,'柔印版2' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_184,'') as 編號,'柔印版3' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_185,'') as 編號,'柔印版4' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_186,'') as 編號,'柔印版5' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_187,'') as 編號,'柔印版6' as 自定義位置  from [" + sql.CYDB +"].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0') union all " +
                        "(select FModel,FNumber,isnull(F_135,'') as 油墨,isnull(F_188,'') as 編號,'柔印版7' as 自定義位置  from [" + sql.CYDB + "].[dbo].[T_ICItem] where FNumber like '14.A%' and Fdeleted = '0'))a where a.Fnumber like '%" + txtValue.Text + "%'  order by a.Fmodel";

                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dgv17F.Rows.Add();
                    dgv17F.Rows[n].Cells[0].Value = item["FModel"].ToString();
                    dgv17F.Rows[n].Cells[1].Value = item["FNumber"].ToString();
                    dgv17F.Rows[n].Cells[2].Value = item["油墨"].ToString();
                    dgv17F.Rows[n].Cells[3].Value = item["編號"].ToString();
                    dgv17F.Rows[n].Cells[4].Value = item["自定義位置"].ToString();
                }
            }

            dt.Clear();
        }
    }
}
