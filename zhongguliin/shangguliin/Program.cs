using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace zhongguliin
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> zinzu = new List<string>() { "精", "清", "從" , "心" , "邪" };
            List<string> jaenzu = new List<string>() { "見","溪","羣","疑","群","曉","匣" };

            try
            {
                Console.TreatControlCAsInput = true;

                Workbook wb = new Workbook(@"D:\\shangguliin.xlsx");
                Worksheet ws = wb.Worksheets[0];
                var dt = ws.Cells.ExportDataTable(0, 0, 10000, 19);
                int k = 0;
                while (!string.IsNullOrEmpty(dt.Rows[k][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k][9].ToString()) || 
                    !string.IsNullOrEmpty(dt.Rows[k+1][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k+1][9].ToString()))
                {
                    string denhuvin = dt.Rows[k][4].ToString().Trim() + dt.Rows[k][5].ToString().Trim() + dt.Rows[k][6].ToString().Trim();
                    if (zinzu.Contains(dt.Rows[k][3].ToString().Trim()))
                    {
                        if (denhuvin == "三合諄" || denhuvin == "一合魂")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "文")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "文";
                            }
                        }
                        else if (denhuvin == "三合仙A" || denhuvin == "二開山" || denhuvin == "二合山")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "仙")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "仙";
                            }
                        }
                    }
                    else if (jaenzu.Contains(dt.Rows[k][3].ToString().Trim()))
                    {
                        if (denhuvin == "一開豪" || denhuvin == "三開幽" || denhuvin == "二開肴")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "幽")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "幽";
                            }
                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "o");
                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ɔ", "o");
                        }
                        else if (denhuvin == "三開尤")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "缶")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "缶";
                            }
                        }
                        else if (denhuvin == "一開侯")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "侯")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "侯";
                            }
                        }
                    }
                    k++;
                }
                wb.Save(@"D:\\shangguliin.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}


