using Aspose.Cells;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System;
using System.Reflection;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace zhongguliin
{
    class Program
    {
        static void Main(string[] args)
        {
            //List<string> zinzu = new List<string>() { "精", "清", "從", "心", "邪" };
            //List<string> jaenzu = new List<string>() { "見", "溪", "羣", "疑", "群", "曉", "匣", "影" };
            try
            {
                gä3shen1diao4();
                //Console.TreatControlCAsInput = true;

                //Workbook wb = new Workbook(@"D:\MyDocument\test1.xlsx");

                ////Workbook wb = new Workbook(@"D:\\shangguliin.xlsx");從韻部到字
                //Worksheet ws = wb.Worksheets[0];
                //var dt = ws.Cells.ExportDataTable(0, 0, 10000, 19);
                //int border = 9720; //最後一個已考行數

                //for (int k = border; !string.IsNullOrEmpty(dt.Rows[k][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k][9].ToString()); k++)
                //{
                //    int i = 0;
                //    for (i = 0; (dt.Rows[i][8].ToString() != dt.Rows[k][8].ToString() || ws.Cells["C" + (i + 1).ToString()].Value == null) && i <= border; i++)
                //    {
                //    }
                //    if (i < border)
                //    {
                //        string apend = ws.Cells["C" + (k + 1).ToString()].Value != null &&
                //            ws.Cells["C" + (k + 1).ToString()].Value.ToString() != ws.Cells["C" + (i + 1).ToString()].Value.ToString() ?
                //            " OLD " + ws.Cells["C" + (k + 1).ToString()].Value.ToString() : "";
                //        ws.Cells["C" + (k + 1).ToString()].Value = ws.Cells["C" + (i + 1).ToString()].Value.ToString() + apend;
                //        ws.Cells["A" + (k + 1).ToString()].Value = ws.Cells["A" + (i + 1).ToString()].Value?.ToString();
                //    }
                //    else if (ws.Cells["C" + (k + 1).ToString()].Value != null)
                //    {
                //        ws.Cells["C" + (k + 1).ToString()].Value = "OLD " + ws.Cells["C" + (k + 1).ToString()].Value.ToString();

                //    }
                //}
                ///*
                //                int k = 0;
                //                while (!string.IsNullOrEmpty(dt.Rows[k][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k][9].ToString()) ||
                //                    !string.IsNullOrEmpty(dt.Rows[k + 1][3].ToString()) || !string.IsNullOrEmpty(dt.Rows[k + 1][9].ToString()))
                //                {
                //                    string denhuvin = dt.Rows[k][4].ToString().Trim() + dt.Rows[k][5].ToString().Trim() + dt.Rows[k][6].ToString().Trim();
                //                    dt.Rows[k][8].ToString().Trim()
                //                    //UnmergeDoInZy(dt, k, ws);
                //                    //ErDaoSang(dt, k, ws)
                //                    //Iou2BiO( denhuvin,  dt,  k, ws);
                //                    /*

                //                    bool isShang = dt.Rows[k][7].ToString().Trim() == "上";

                //                    if (dt.Rows[k][3].ToString().Trim() == "影")
                //                    {   
                //                        ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ʔ", "h") + (isShang? "ɣɣ" : "");
                //                        ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ʔ", "h") + (isShang ? "ɣɣ" : "");
                //                    }

                //                    if (denhuvin == "三合文")
                //                    {
                //                        ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ʷə", "o");
                //                        ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ʷə", "o");

                //                        ws.Cells["A" + (k + 1).ToString()].Value = "文";
                //                    }
                //                    else if (denhuvin == "三合諄" || denhuvin == "一合魂")
                //                    {

                //                        if (dt.Rows[k][0].ToString().Trim() != "諄")
                //                        {
                //                            ws.Cells["A" + (k + 1).ToString()].Value = "諄";
                //                        }
                //                    }
                //                    if (zinzu.Contains(dt.Rows[k][3].ToString().Trim()))
                //                    {
                //                        if (denhuvin == "三合諄" || denhuvin == "一合魂")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "文")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "文";
                //                            }
                //                        }
                //                        else if (denhuvin == "三合仙A" || denhuvin == "二開山" || denhuvin == "二合山")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "仙")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "仙";
                //                            }
                //                        }
                //                    }
                //                    else if (jaenzu.Contains(dt.Rows[k][3].ToString().Trim()))
                //                    {
                //                        if (denhuvin == "一開豪" || denhuvin == "三開幽" || denhuvin == "二開肴")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "幽")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "幽";
                //                            }
                //                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "o");
                //                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ɔ", "o");
                //                        }
                //                        else if (denhuvin == "三開屋" || denhuvin == "一開沃" || denhuvin == "二開覺")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "覺")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "覺";
                //                            }
                //                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "o");
                //                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("ɔ", "o");
                //                        }
                //                        else if (denhuvin == "三開燭" || denhuvin == "一開屋")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "屋")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "屋";
                //                            }
                //                        }
                //                        //else if (denhuvin == "三開職" || denhuvin == "二合麥")
                //                        //{
                //                        //    if (dt.Rows[k][0].ToString().Trim() != "職")
                //                        //    {
                //                        //        ws.Cells["A" + (k + 1).ToString()].Value = "職";
                //                        //    }
                //                        //}
                //                        else if (denhuvin == "三開尤" )
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "缶")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "缶";
                //                            }
                //                            ws.Cells["B" + (k + 1).ToString()].Value = dt.Rows[k][1].ToString().Trim().Replace("ɔ", "ɤ");
                //                            ws.Cells["C" + (k + 1).ToString()].Value = dt.Rows[k][2].ToString().Trim().Replace("o", "ɤ");
                //                        }
                //                        else if (denhuvin == "三開魚")
                //                        {
                //                            if (dt.Rows[k][0].ToString().Trim() != "侯")
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = "侯";
                //                            }
                //                        }
                //                        else if (denhuvin == "一開侯")
                //                        {
                //                            var uen = "侯";
                //                            if (dt.Rows[k][1].ToString().Trim().EndsWith("oks"))
                //                                uen = "覺";
                //                            else if (dt.Rows[k][1].ToString().Trim().EndsWith("ɔks"))
                //                                uen = "屋";
                //                            if (dt.Rows[k][0].ToString().Trim() != uen)
                //                            {
                //                                ws.Cells["A" + (k + 1).ToString()].Value = uen;
                //                            }
                //                        }
                //                    }
                //                    *
                //                    k++;
                //                }*/
                //wb.Save(@"D:\MyDocument\shangguliin3.xlsx");
                ////wb.Save(@"D:\\shangguliin3.xlsx");
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

        private static void Gi3Shang4Gu3Biao3Ja1Pin1In1()
        {
            Workbook wb = new Workbook(@"D:\MyDocument\test1.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9720, 9);
            Workbook wb2 = new Workbook(@"D:\MyDocument\test2.xlsx");//中古南京表
            Worksheet ws2 = wb2.Worksheets[0];
            var dt2 = ws2.Cells.ExportDataTable(0, 0, 3885, 9);
            for (int k = 0; k < 240; k++)
            {                
                var fan3cie5 = dt.Rows[k][5].ToString().Trim();
                var sdhüd = dt.Rows[k][0].ToString().Trim() + dt.Rows[k][1].ToString().Trim() + dt.Rows[k][2].ToString().Trim() + dt.Rows[k][3].ToString().Trim() + dt.Rows[k][4].ToString().Trim();
                int i = 0;
                while (dt2.Rows[i][5].ToString().Trim() != fan3cie5 && i < 3884)
                {
                    i++;
                }
                if (i < 3884 || (dt2.Rows[i][5].ToString().Trim() == fan3cie5))
                {//找到
                    ws.Cells["H" + (k + 1).ToString()].Value = dt2.Rows[i][7].ToString().Trim();
                    ws.Cells["I" + (k + 1).ToString()].Value = dt2.Rows[i][8].ToString().Trim();
                }
                else
                {
                    i = 0;
                    while (dt2.Rows[i][0].ToString().Trim() + dt2.Rows[i][1].ToString().Trim() + dt2.Rows[i][2].ToString().Trim() + dt2.Rows[i][3].ToString().Trim() + dt2.Rows[i][4].ToString().Trim() != sdhüd && i < 3884)
                    {
                        i++;
                        if (i == 1618)
                        {
                            int asjdk = 0;
                        }
                    }
                    if (i < 3884 || (dt2.Rows[i][0].ToString().Trim() + dt2.Rows[i][1].ToString().Trim() + dt2.Rows[i][2].ToString().Trim() + dt2.Rows[i][3].ToString().Trim() + dt2.Rows[i][4].ToString().Trim() == sdhüd))
                    {//找到
                        ws.Cells["H" + (k + 1).ToString()].Value = dt2.Rows[i][7].ToString().Trim();
                        ws.Cells["I" + (k + 1).ToString()].Value = dt2.Rows[i][8].ToString().Trim();
                    }
                }
            }
            wb.Save(@"D:\MyDocument\shangguliin.xlsx");
        }

        private static void Fen1kä1ho5bin4gä5()
        {
            Workbook wb = new Workbook(@"D:\1.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9730, 1);
            int ii;
            for (int k = 0;k<9730 ; k++)
            {
                if (ws.Cells["A" + (k + 1).ToString()].IsMerged)
                { /*
                    if (ws.Cells["A" + (k + 1).ToString()].Value?.ToString() != ws.Cells["B" + (k + 1).ToString()].Value?.ToString())
                        ii = 0;*/
                    if (ws.Cells["A" + (k + 1).ToString()].Value != null)
                    {
                        int i = 1;
                        SetColor(ws, "A" + (k + 1).ToString(), Color.Red);
                        while (ws.Cells["A" + (k + 1 +i).ToString()].Value == null)
                        {
                            ws.Cells["A" + (k + 1 + i).ToString()].Value = ws.Cells["A" + (k + 1).ToString()].Value.ToString().Trim();
                            SetColor(ws, "A" + (k + 1 +i).ToString(), Color.Red);
                            i++;
                        }
                        ws.Cells.UnMerge(k, 0, i, 1);

                        k += i-1;
                    }

                }
            }
            wb.Save(@"D:\shangguliin.xlsx");
        }

        private static void SetColor(Worksheet worksheet, string coord, Color c)
        {
            Style style = new Style();
            style.Font.Color=c;
            worksheet.Cells[coord].SetStyle(style);
        }

        private static void cä5gä5()
        {
            Workbook wb = new Workbook(@"D:\1.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            for (int k = 0; k < 9730; k++)
            {
                if (ws.Cells["A" + (k + 1).ToString()].IsMerged)
                {                   
                    ws.Cells.UnMerge(k, 0, 2, 1);
                }
            }
            wb.Save(@"D:\shangguliin.xlsx");
        }

        private static void gä3shen1diao4()
        {   //如果表中I列是上、去聲，而D列沒ɣ、h尾，說明上古音有誤，要改。
            
            Workbook wb = new Workbook(@"D:\shang4gu3li3in1.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            for (int k = 0; k < 9919; k++)
            {
                if (ws.Cells["D" + (k + 1).ToString()].Value != null)
                {
                    if (ws.Cells["I" + (k + 1).ToString()].Value.ToString() == "上" && !ws.Cells["D" + (k + 1).ToString()].Value.ToString().EndsWith("ɣ"))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString() + "ɣ";
                    }
                    else if (ws.Cells["I" + (k + 1).ToString()].Value.ToString() == "去" && !ws.Cells["D" + (k + 1).ToString()].Value.ToString().EndsWith("h"))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString() + "h";
                    }
                    else if (ws.Cells["I" + (k + 1).ToString()].Value.ToString() == "平" && 
                        (ws.Cells["D" + (k + 1).ToString()].Value.ToString().EndsWith("h") || ws.Cells["D" + (k + 1).ToString()].Value.ToString().EndsWith("ɣ")))
                    {
                        string guin = ws.Cells["D" + (k + 1).ToString()].Value.ToString();
                        ws.Cells["D" + (k + 1).ToString()].Value = guin.Substring(0, guin.Length-1) ;
                    }
                    if (ws.Cells["F" + (k + 1).ToString()].Value.ToString() == "三")
                    {//如果表中F列是三等，而D列有ˤ，說明上古音有誤，要改。
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ˤ", "");
                    }
                    else if (!ws.Cells["D" + (k + 1).ToString()].Value.ToString().Contains("ˤ"))
                    {//如果表中F列是丰等，而D列無ˤ，說明上古音有誤，要改。
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("a", "ˤa").Replace("ɔ", "ˤɔ").Replace("o", "ˤo").Replace("ə", "ˤə").Replace("e", "ˤe").Replace("ɛ", "ˤɛ");
                    }

                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "咸")
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("am", "əm");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "嚴")
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("am", "om");
                    }
                    //if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "脂A" && ws.Cells["E" + (k + 1).ToString()].Value.ToString() == "曉")
                    if ("脂A質A".Contains(ws.Cells["H" + (k + 1).ToString()].Value.ToString()))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ə", "e");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "尤" || ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "冬")
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ɔ", "o");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "質A")
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ɔ", "et");
                    }
                    if ("肴麻".Contains(ws.Cells["H" + (k + 1).ToString()].Value.ToString()) && ws.Cells["F" + (k + 1).ToString()].Value.ToString() == "二"  && !ws.Cells["D" + (k + 1).ToString()].Value.ToString().Contains("r"))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ˤ", "rˤ");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString()== "麥" && ws.Cells["F" + (k + 1).ToString()].Value.ToString() == "二" &&
                        "見群溪疑曉匣".Contains(ws.Cells["E" + (k + 1).ToString()].Value.ToString()))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ɛ", "ə");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "真A" &&
                        "見群溪疑曉匣".Contains(ws.Cells["E" + (k + 1).ToString()].Value.ToString()))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("on", "ʷen");
                    }
                    if ("文物".Contains(ws.Cells["H" + (k + 1).ToString()].Value.ToString()) &&
                        "見群溪疑曉匣".Contains(ws.Cells["E" + (k + 1).ToString()].Value.ToString()))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ʷə", "o");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "物" &&
                        "幫滂並明".Contains(ws.Cells["E" + (k + 1).ToString()].Value.ToString()))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("o", "ə");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "魚" &&
                       ws.Cells["E" + (k + 1).ToString()].Value.ToString() == "心")
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("xa", "xʷa").Replace("ŋa", "ŋʷa");
                    }
                   /* if ("溪滂透".Contains(ws.Cells["E" + (k + 1).ToString()].Value.ToString()) &&
                        !ws.Cells["D" + (k + 1).ToString()].Value.ToString().Contains("ʰ"))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("k", "kʰ").Replace("p", "pʰ").Replace("t", "tʰ");
                    }*/
                }
            }
            wb.Save(@"D:\shangguliin.xlsx");
        }
    }
}


