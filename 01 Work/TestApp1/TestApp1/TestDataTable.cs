using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TestApp1
{
    class TestDataTable
    {
        public DataTable WriteDT(DataTable dt)
        {
            // 加列 
            dt.Columns.Add("序号", Type.GetType("System.Int32"));
            dt.Columns.Add("项目工号", Type.GetType("System.String"));
            dt.Columns.Add("预留工号", Type.GetType("System.String"));
            dt.Columns.Add("合同签订方式", Type.GetType("System.String"));
            dt.Columns.Add("项目负责人", Type.GetType("System.String"));
            dt.Columns.Add("年度", Type.GetType("System.String"));
            dt.Columns.Add("到款", Type.GetType("System.String"));
            dt.Columns.Add("小计", Type.GetType("System.String"));
            dt.Columns.Add("材料费", Type.GetType("System.String"));
            dt.Columns.Add("专用费", Type.GetType("System.String"));
            dt.Columns.Add("外协费", Type.GetType("System.String"));
            dt.Columns.Add("工作量的支撑工时（小时）", Type.GetType("System.String"));
            dt.Columns.Add("备注", Type.GetType("System.String"));


            // 加行 
            DataRow dr = dt.NewRow();
            dr["序号"] = 1;
            dr["项目工号"] = "Q176-2018-Y369";
            dr["预留工号"] = "Q176-2018-Y369";
            dr["合同签订方式"] = "电子";
            dr["项目负责人"] = "tl";
            dr["年度"] = "合同约定";
            dr["备注"] = "cl";
            dt.Rows.Add(dr);

            DataRow dr1 = dt.NewRow();
            dr1["年度"] = "全周期预算合计";
            dr1["到款"] = "150";
            dr1["小计"] = "150";
            dr1["材料费"] = "150";
            dr1["专用费"] = "150";
            dr1["外协费"] = "150";
            dr1["工作量的支撑工时（小时）"] = "150";
            dr1["备注"] = "cl";
            dt.Rows.Add(dr1);

            DataRow dr2 = dt.NewRow();
            dr2["年度"] = "2018";
            dr2["到款"] = "10";
            dr2["小计"] = "10";
            dr2["材料费"] = "10";
            dr2["专用费"] = "10";
            dr2["外协费"] = "10";
            dr2["工作量的支撑工时（小时）"] = "";
            dr2["备注"] = "cl";
            dt.Rows.Add(dr2);

            DataRow dr3 = dt.NewRow();
            dr3["年度"] = "2019";
            dr3["到款"] = "20";
            dr3["小计"] = "20";
            dr3["材料费"] = "20";
            dr3["专用费"] = "20";
            dr3["外协费"] = "20";
            dr3["工作量的支撑工时（小时）"] = "";
            dr3["备注"] = "cl";
            dt.Rows.Add(dr3);




            DataRow dr4 = dt.NewRow();
            dr4["序号"] = 2;
            dr4["项目工号"] = "Q178-2020-Y369";
            dr4["预留工号"] = "Q178-2020-Y369";
            dr4["合同签订方式"] = "电子";
            dr4["项目负责人"] = "tl";
            dr4["年度"] = "合同约定";
            dr4["备注"] = "cl";
            dt.Rows.Add(dr4);

            DataRow dr5 = dt.NewRow();
            dr5["年度"] = "全周期预算合计";
            dr5["到款"] = "100";
            dr5["小计"] = "100";
            dr5["材料费"] = "100";
            dr5["专用费"] = "100";
            dr5["外协费"] = "100";
            dr5["工作量的支撑工时（小时）"] = "100";
            dr5["备注"] = "cl";
            dt.Rows.Add(dr5);

            DataRow dr6 = dt.NewRow();
            dr6["年度"] = "2018";
            dr6["到款"] = "10";
            dr6["小计"] = "10";
            dr6["材料费"] = "10";
            dr6["专用费"] = "10";
            dr6["外协费"] = "10";
            dr6["工作量的支撑工时（小时）"] = "";
            dr6["备注"] = "cl";
            dt.Rows.Add(dr6);

            DataRow dr7 = dt.NewRow();
            dr7["年度"] = "2019";
            dr7["到款"] = "20";
            dr7["小计"] = "20";
            dr7["材料费"] = "20";
            dr7["专用费"] = "20";
            dr7["外协费"] = "20";
            dr7["工作量的支撑工时（小时）"] = "";
            dr7["备注"] = "cl";
            dt.Rows.Add(dr7);

            return dt;
        }


        public DataTable exportDT(DataTable dt)
        {
            // 1. 删除不需要的列
            // 2. 项目工号为空的，与前一行的项目工号相同
            // 3. 年度不在1900-2099之间的行，都删除
            string[] strArray = new string[] { "项目工号","年度","到款","小计","材料费","专用费","外协费", "工作量的支撑工时（小时）", "备注"};

            int colCount = dt.Columns.Count;
            for (int i = colCount - 1; i >= 0; i--)
            {
                DataColumn col = dt.Columns[i];
   
                if (Array.IndexOf<string>(strArray, col.ColumnName) == -1)
                {
                    dt.Columns.Remove(col);
                }
            }

            int rowCount = dt.Rows.Count;
            for (int i = 0; i < rowCount; i++)
            {
                DataRow row = dt.Rows[i];
                string xmgh = row["项目工号"].ToString();

                if (xmgh == "")
                {
                    string prexmgh = dt.Rows[i - 1]["项目工号"].ToString();
                    row["项目工号"] = prexmgh;
                }
               
            }

            for (int i = rowCount - 1; i >= 0; i--)
            {
                DataRow row = dt.Rows[i];
                string year = row["年度"].ToString();
                if (!IsNumeric(year))
                {
                    dt.Rows.Remove(row);
                }
            }

            return dt;
        }

        public static bool IsNumeric(string value)
        {
            // 1900-2099年
            return Regex.IsMatch(value, @"^(19|20)\d{2}$");
        }
    }
}
