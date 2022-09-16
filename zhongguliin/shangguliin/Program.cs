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
            List<string> jaenzu = new List<string>() { "見","溪","羣","疑","群","曉","匣", "影"};

            try
            {
                Console.TreatControlCAsInput = true;

                Workbook wb = new Workbook(@"C:\\Users\yli\OneDrive - Senacor Technologies AG\Dokumente\shangguliin.xlsx");
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
                        else if (denhuvin == "三開屋" || denhuvin == "一開沃" || denhuvin == "二開覺")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "覺")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "覺";
                            }
                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "o");
                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ɔ", "o");
                        }
                        else if (denhuvin == "三開燭" || denhuvin == "一開屋")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "屋")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "屋";
                            }
                        }
                        //else if (denhuvin == "三開職" || denhuvin == "二合麥")
                        //{
                        //    if (dt.Rows[k][0].ToString().Trim() != "職")
                        //    {
                        //        ws.Cells["A" + (k + 1).ToString()].Value = "職";
                        //    }
                        //}
                        else if (denhuvin == "三開尤" )
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "缶")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "缶";
                            }
                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "ɤ");
                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("o", "ɤ");
                        }
                        else if (denhuvin == "三開魚")
                        {
                            if (dt.Rows[k][0].ToString().Trim() != "侯")
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = "侯";
                            }
                        }
                        else if (denhuvin == "一開侯")
                        {
                            var uen = "侯";
                            if (dt.Rows[k][1].ToString().Trim().EndsWith("oks"))
                                uen = "覺";
                            else if (dt.Rows[k][1].ToString().Trim().EndsWith("ɔks"))
                                uen = "屋";
                            if (dt.Rows[k][0].ToString().Trim() != uen)
                            {
                                ws.Cells["A" + (k + 1).ToString()].Value = uen;
                            }
                        }
                    }
                    k++;
                }
                wb.Save(@"C:\\Users\yli\OneDrive - Senacor Technologies AG\Dokumente\shangguliin.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}


