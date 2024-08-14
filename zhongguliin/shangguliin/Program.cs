﻿using Aspose.Cells;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System;

namespace zhongguliin
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> zinzu = new List<string>() { "精", "清", "從", "心", "邪" };
            List<string> jaenzu = new List<string>() { "見", "溪", "羣", "疑", "群", "曉", "匣", "影" };
            try
            {
                Console.TreatControlCAsInput = true;

                Workbook wb = new Workbook(@"D:\MyDocument\test1.xlsx");

                //Workbook wb = new Workbook(@"D:\\shangguliin.xlsx");從韻部到字
                Worksheet ws = wb.Worksheets[0];
                var dt = ws.Cells.ExportDataTable(0, 0, 10000, 19);
                int border=9540; //最後一個已考行數

                for (int k = border; !string.IsNullOrEmpty(dt.Rows[k][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k][9].ToString()); k++)
                {
                    int i = 0;
                    for (i = 0; (dt.Rows[i][8].ToString() != dt.Rows[k][8].ToString() || ws.Cells["C" + (i + 1).ToString()].Value == null) && i <= border ; i++)
                    {
                    }
                    if (i<border)
                    {
                        string apend = ws.Cells["C" + (k + 1).ToString()].Value != null && 
                            ws.Cells["C" + (k + 1).ToString()].Value.ToString() != ws.Cells["C" + (i + 1).ToString()].Value.ToString() ? 
                            " OLD " + ws.Cells["C" + (k + 1).ToString()].Value.ToString() : "";
                        ws.Cells["C" + (k + 1).ToString()].Value = ws.Cells["C" + (i + 1).ToString()].Value.ToString() + apend;
                        ws.Cells["A" + (k + 1).ToString()].Value = ws.Cells["A" + (i + 1).ToString()].Value?.ToString();
                    }
                    else if (ws.Cells["C" + (k + 1).ToString()].Value!=null)
                    {
                        ws.Cells["C" + (k + 1).ToString()].Value = "OLD " + ws.Cells["C" + (k + 1).ToString()].Value.ToString();

                    }
                }
                /*
                                int k = 0;
                                while (!string.IsNullOrEmpty(dt.Rows[k][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k][9].ToString()) ||
                                    !string.IsNullOrEmpty(dt.Rows[k + 1][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k + 1][9].ToString()))
                                {
                                    string denhuvin = dt.Rows[k][4].ToString().Trim() + dt.Rows[k][5].ToString().Trim() + dt.Rows[k][6].ToString().Trim();
                                    dt.Rows[k][8].ToString().Trim()
                                    //UnmergeDoInZy(dt, k, ws);
                                    //ErDaoSang(dt, k, ws)
                                    //Iou2BiO( denhuvin,  dt,  k, ws);
                                    /*

                                    bool isShang = dt.Rows[k][7].ToString().Trim() == "上";
                                    
                                    if (dt.Rows[k][3].ToString().Trim() == "影")
                                    {   
                                        ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ʔ", "h") + (isShang? "ɣɣ" : "");
                                        ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ʔ", "h") + (isShang ? "ɣɣ" : "");
                                    }

                                    if (denhuvin == "三合文")
                                    {
                                        ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ʷə", "o");
                                        ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ʷə", "o");

                                        ws.Cells["A" + (k + 1).ToString()].Value = "文";
                                    }
                                    else if (denhuvin == "三合諄" || denhuvin == "一合魂")
                                    {

                                        if (dt.Rows[k][0].ToString().Trim() != "諄")
                                        {
                                            ws.Cells["A" + (k + 1).ToString()].Value = "諄";
                                        }
                                    }
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
                                    *
                                    k++;
                                }*/
                wb.Save(@"D:\MyDocument\shangguliin3.xlsx");
                //wb.Save(@"D:\\shangguliin3.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void UnmergeDoInZy(DataTable dt, int k, Worksheet ws)
        {
            if (dt.Rows[k][9].ToString().Trim() == "")
            {
                var range = ws.Cells["J" + (k + 1).ToString()].GetMergedRange();
                if (range != null)
                {
                    string[] diihodier = range.Address.Split(":");
                    int diisu = Convert.ToInt32(diihodier[0].Substring(1));
                    int diersu = Convert.ToInt32(diihodier[1].Substring(1));
                    string zy = "";
                    for (int j = diisu; j < diersu; j++)
                    {
                        if (dt.Rows[j - 1][9].ToString().Trim() != "")
                        {
                            zy = dt.Rows[j - 1][9].ToString().Trim();
                            break;
                        }
                    }
                    ws.Cells.UnMerge(diisu - 1, 9, diersu - diisu + 1, 1);
                    ws.Cells["J" + (k + 1).ToString()].Value = zy;
                }
            }
        }

        private static void ErDaoSang(DataTable dt, int k, Worksheet ws)
        {
            //第二列到第三列
            if (dt.Rows[k][2].ToString().Trim() == "" && dt.Rows[k][1].ToString().Trim() != "")
            {
                ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim();
                ws.Cells["B" + (k + 1).ToString()].Value = "";
            }
        }

        private static void Iou2BiO(string denhuvin, DataTable dt, int k, Worksheet ws)
        {
            if (denhuvin == "三開尤")
            {
                ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ə", "o");
                ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ə", "o");
                ws.Cells["A" + (k + 1).ToString()].Value = "幽";
            }
        }
    }
}


