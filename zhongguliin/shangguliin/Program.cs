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
                Si5Mä5();
                //Hao2Siao1();
                //gae3chen2guae5zi5in1biao1();
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
            for (int k = 0; k < 9730; k++)
            {
                if (ws.Cells["A" + (k + 1).ToString()].IsMerged)
                { /*
                    if (ws.Cells["A" + (k + 1).ToString()].Value?.ToString() != ws.Cells["B" + (k + 1).ToString()].Value?.ToString())
                        ii = 0;*/
                    if (ws.Cells["A" + (k + 1).ToString()].Value != null)
                    {
                        int i = 1;
                        SetColor(ws, "A" + (k + 1).ToString(), Color.Red);
                        while (ws.Cells["A" + (k + 1 + i).ToString()].Value == null)
                        {
                            ws.Cells["A" + (k + 1 + i).ToString()].Value = ws.Cells["A" + (k + 1).ToString()].Value.ToString().Trim();
                            SetColor(ws, "A" + (k + 1 + i).ToString(), Color.Red);
                            i++;
                        }
                        ws.Cells.UnMerge(k, 0, i, 1);

                        k += i - 1;
                    }

                }
            }
            wb.Save(@"D:\shangguliin.xlsx");
        }

        private static void SetColor(Worksheet worksheet, string coord, Color c)
        {
            Style style = new Style();
            style.Font.Color = c;
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
                        ws.Cells["D" + (k + 1).ToString()].Value = guin.Substring(0, guin.Length - 1);
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
                    if ("肴麻".Contains(ws.Cells["H" + (k + 1).ToString()].Value.ToString()) && ws.Cells["F" + (k + 1).ToString()].Value.ToString() == "二" && !ws.Cells["D" + (k + 1).ToString()].Value.ToString().Contains("r"))
                    {
                        ws.Cells["D" + (k + 1).ToString()].Value = ws.Cells["D" + (k + 1).ToString()].Value.ToString().Replace("ˤ", "rˤ");
                    }
                    if (ws.Cells["H" + (k + 1).ToString()].Value.ToString() == "麥" && ws.Cells["F" + (k + 1).ToString()].Value.ToString() == "二" &&
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
        /// <summary>
        /// 給 https://github.com/osfans/MCPDict/blob/master/tools/tables/data/%E5%BB%A3%E9%9F%BB.tsv 加廣通羅馬字
        /// </summary>
        private static void ja1guang3tong1()
        {
            Workbook wb = new Workbook(@"D:\old.xlsx");//舊
            Worksheet ws = wb.Worksheets[0];
            Workbook wb2 = new Workbook(@"D:\Guangyun_Langjin_Zhonggu.1.0.xlsx");//舊
            Worksheet ws2 = wb2.Worksheets[0];
            for (int k = 2; k <= 25332; k++)
            {
                var fang3cie5 = ws.Cells["H" + k.ToString()].Value.ToString().Replace("式之(脂)", "式脂").Replace("叉⟨尺⟩隹", "尺隹").Replace("居帋", "居氏").Replace("都搕⟨榼⟩", "都榼")
                    .Replace("徂累(壘)", "徂累").Replace("近⟨丘⟩倨", "丘倨").Replace("博故", "愽故").Replace("臧𧙓⟨祚⟩", "臧祚").Replace("之芮", "之銳").Replace("土⟨士⟩列", "士列")
                    .Replace("陟離", "陟移").Replace("姊宜⦉規⦊", "姊規").Replace("子⦅?⦆⟨之⟩垂", "專垂").Replace("側宜", "側移").Replace("士宜", "士移").Replace("杜懷", "柱懷")
                    .Replace("乙皆(乖)", "乙乖").Replace("諧⟨諾⟩皆", "諾皆").Replace("昌來⟨求⟩", "昌求").Replace("普才⦅來⦆⟨求⟩", "普求").Replace("側(職)鄰", "職鄰").Replace("府(撫)文", "撫文")
                    .Replace("嘗芮", "甞芮").Replace("呼吠", "呼吠？").Replace("他⟨迍⟩怪", "迍怪").Replace("古賣(邁)", "古邁").Replace("五夾⦅洽⦆⟨冷⟩", "五剄")
                    .Replace("力頑⟨規⟩", "力頑").Replace("崇⟨?⟩玄⟨?⟩", "崇玄").Replace("居乙⟨乞⟩", "居乞").Replace("都牢", "都勞").Replace("子𩨷", "子𩩆").Replace("縷𩨷", "縷𩨭")
                    .Replace("千侯⟨隹⟩", "尺隹").Replace("子幽⟨絲⟩", "子之").Replace("山幽⟨函⟩", "蘇含").Replace("昨⟨作⟩三", "作三").Replace("符咸(䒦)", "符䒦")
                    .Replace("呂張", "呂章").Replace("居理", "居里").Replace("初紀⦅己⦆⟨乙⟩", "初栗").Replace("美畢(筆)", "美筆").Replace("居竭(謁)", "居竭")
                    .Replace("求⟨?⟩蟹", "求蟹").Replace("求⟨乖⟩蟹", "求蟹").Replace("而允⟨兗⟩", "而兖").Replace("許竭(謁)", "許竭").Replace("五骨⟨滑⟩", "五滑")
                    .Replace("辝纂⦅短⦆⟨矩⟩", "似矩").Replace("被⟨披⟩免", "披免").Replace("烏晈", "烏皎").Replace("以沼⦅小⦆⟨水⟩", "以水")
                    .Replace("作⦅子⦆⟨千⟩可", "千可").Replace("博下", "愽下").Replace("博蓋", "愽蓋").Replace("莫幸(杏)", "莫杏").Replace("呼⟨乎⟩䁝", "乎䁝")
                    .Replace("苦蓋(愛)", "苦愛").Replace("方⟨芳⟩廢", "芳廢").Replace("徐(疾)刃", "疾刃").Replace("芳⦅反⦆⟨叉⟩万", "初万").Replace("祖⟨徂⟩贊", "徂贊").Replace("姝⟨殊⟩雪", "殊雪")
                    .Replace("式任⟨荏⟩", "式荏").Replace("魯⟨魚⟩掩⟨埯⟩", "魚埯").Replace("子⟨千⟩仲", "千仲").Replace("乙冀", "乙兾").Replace("扶涕⟨沸⟩", "扶沸")
                    .Replace("博慢⟨漫⟩", "博漫").Replace("于⟨予⟩線", "予線").Replace("于⟨子⟩亮", "子亮").Replace("許⦅火⦆令⟨含⟩", "火含").Replace("[徒]候", "徒候").Replace("丘謁(竭)", "丘竭")
                    .Replace("音黯去聲", "乙鑒").Replace("音蒸上聲", "之庱").Replace("矛⟨予⟩割", "予割").Replace("簪⦅子⦆⟨于⟩摑", "胡麥").Replace("居列(?)", "居列").Replace("廁列⟨別⟩", "厠別")
                #region 柳漫另有說法
                    .Replace("於力(棘)", "於力").Replace("倉雜(臘)", "倉雜").Replace("七⟨火⟩役", "七役").Replace("之⦅志⦆⟨?⟩役", "之役").Replace("士⦅仕⦆⟨丘⟩㾕", "士㾕").Replace("香⦅許⦆幽(彪)", "香幽")
                    .Replace("九輦(善)", "九輦").Replace("毗養⦅兩⦆⟨霄⟩", "毗養").Replace("士⟨于⟩忍", "七忍").Replace("丘之⟨乏⟩", "丘之").Replace("丑戾⟨居⟩", "丑戾").Replace("下珍⟨殄⟩", "下珍")
                    .Replace("初紀⦅史⦆⟨夬⟩", "初紀").Replace("火⟨丈⟩弔⦅叫⦆⟨列⟩", "火弔").Replace("火弔⟨即⟩", "火弔").Replace("丈⟨?⟩夥⟨黠⟩", "丈夥").Replace("花⟨?⟩夥⟨黠⟩", "花夥")
                    .Replace("職⦉且⦊勇", "職勇").Replace("丁全⟨兮⟩", "跪頑")
                #endregion
                    .Replace("士⟨七⟩", "七").Replace("七⟨士⟩", "士")
                    .Replace("戶", "戸").Replace("真", "眞").Replace("菹", "葅").Replace("暨", "曁").Replace("顛", "顚").Replace("彥", "彦").Replace("鑒", "鑑")
                    .Replace("既", "旣").Replace("溉", "漑").Replace("槩", "概").Replace("教", "敎").Replace("亙", "亘").Replace("劒", "劔").Replace("𧸖", "賺").Replace("祿", "禄")
                    .Replace("顏", "顔").Replace("虛", "虚").Replace("㚷", "妳").Replace("囂", "嚻").Replace("乘", "乗").Replace("恆", "恒").Replace("摠", "揔").Replace("錄", "録")
                    .Replace("毀", "毁").Replace("豨", "狶").Replace("疎", "踈").Replace("袞", "衮").Replace("眾", "衆").Replace("弊", "獘").Replace("內", "内").Replace("兗", "兖")
                    .Replace("沒", "没").Replace("查", "査");
                //.Replace("七溜⦅霤⦆⟨雷⟩", "七溜")
                int i = 2;
                while (ws2.Cells["I" + i.ToString()].Value?.ToString() != fang3cie5 && i < 3887)
                {
                    i++;
                }
                if (i <= 3887 && ws2.Cells["I" + i.ToString()].Value.ToString() == fang3cie5)
                    ws.Cells["M" + k.ToString()].Value = ws2.Cells["K" + i.ToString()].Value.ToString();
                else
                    ws.Cells["M" + k.ToString()].Value += "Wrong!";
            }
            wb.Save(@"D:\new.xlsx");
        }

        private static void gae3chen2guae5zi5in1biao1()
        {
            StreamReader sr = new StreamReader("D:\\aa.yml");
            StreamWriter sw = new StreamWriter("D:\\南京字表.IPA.yml");
            //Read the first line of text
            var line = sr.ReadLine();
            //Continue to read until you reach end of file
            while (line != null)
            {
                //write the line to console window
                var ipa=line.Replace("5", "ʔ‹7›").Replace("1", "‹1›").Replace("2", "‹2›").Replace("3", "‹3›").Replace("4", "‹5›")
                    .Replace("zhr", "tʂʅ").Replace("chr", "tʂʰʅ").Replace("shr", "ʂʅ").Replace("rʔ‹7›", "ʐʅʔ‹7›")
                    .Replace("ao", "ɔo").Replace("ou", "əɯ").Replace("er", "ɚ").Replace("en", "ən").Replace("ei", "əi").Replace("än", "en").Replace("äʔ", "ɜʔ")
                    .Replace("ng", "ŋ").Replace("zh", "ʈʂ").Replace("ch", "tʂʰ").Replace("sh", "ʂ")
                    .Replace("p", "pʰ").Replace("t", "tʰ").Replace("k", "kʰ").Replace("q", "tɕʰ").Replace("c", "tsʰ")
                    .Replace("b", "p").Replace("d", "t").Replace("g", "k").Replace("j", "tɕ").Replace("z", "ts")                    
                    .Replace("x", "ɕ").Replace("y", "ɿ").Replace("ü", "y").Replace("ä", "ae").Replace("r", "ʐ")
                    .Replace("ʰʂʰ", "ʂʰ").Replace("ʰsʰ", "sʰ").Replace("ʰɕʰ", "ɕʰ");
                sw.WriteLine(ipa);
                //Read the next line
                line = sr.ReadLine();
            }
            //close the file
            sr.Close();
            sw.Close();
        }

        /*
         * ˤø四開蕭；ʷˤø一開豪
         * ˤo一開豪；ʷˤo四開蕭
         */
        private static void Hao2Siao1()
        {
            Workbook wb = new Workbook(@"D:\A.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9913, 1);
            int ii;
            for (int k = 2; k < 9913; k++)
            {
                string zhongguüinmu = ws.Cells["K" + (k + 1).ToString()].Value.ToString();//K列是中古韻母

                if (ws.Cells["G" + (k + 1).ToString()].Value != null)
                {
                    string shangguin = ws.Cells["G" + (k + 1).ToString()].Value.ToString();//G列是上古擬音
                    if (!shangguin.Contains("ˤok") && !shangguin.Contains("ˤoŋ") && !shangguin.Contains("ˤom") && !shangguin.Contains("ˤøk") && !shangguin.Contains("ˤøŋ")
                        && !shangguin.Contains("ˤøl") && !shangguin.Contains("ˤøt") && !shangguin.Contains("ˤøn") && !shangguin.Contains("ˤøm") && !shangguin.Contains("ˤøp"))
                    {
                        if (zhongguüinmu == "豪")
                        {
                            if (shangguin.Contains("ˤo"))
                            {
                                if (shangguin.Contains("ʷˤo"))
                                    ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ʷˤo", "ˤo");
                            }
                            else if (shangguin.Contains("ˤø"))
                            {
                                if (!shangguin.Contains("ʷˤø") && !shangguin.StartsWith("b") && !shangguin.StartsWith("p") && !shangguin.StartsWith("m") && !shangguin.StartsWith("xm"))
                                    ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ˤø", "ʷˤø");
                            }
                        }
                        else if (zhongguüinmu == "蕭")
                        {
                            if (shangguin.Contains("ˤo"))
                            {
                                if (!shangguin.Contains("ʷˤo") && !shangguin.StartsWith("b") && !shangguin.StartsWith("p") && !shangguin.StartsWith("m"))
                                    ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ˤo", "ʷˤo");
                            }
                            else if (shangguin.Contains("ˤø"))
                            {
                                if (shangguin.Contains("ʷˤø"))
                                    ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ʷˤø", "ˤø");
                                else if ((shangguin.StartsWith("s") || shangguin.StartsWith("r") || shangguin.StartsWith("k") || shangguin.StartsWith("g") || shangguin.StartsWith("x") || shangguin.StartsWith("ŋ")) && shangguin.Contains("lˤø"))
                                    ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("lˤø", "ˤø");
                            }
                        }
                    }
                }
            }
            wb.Save(@"D:\shangguliin.xlsx");
            
        }

        private static void Si5Mä5()
        {
            Workbook wb = new Workbook(@"D:\《廣韻》形聲考李.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9912, 1);
            int ii;
            for (int k = 2; k < 9912; k++)
            {
                string zhongguüinmu = ws.Cells["I" + (k + 1).ToString()].Value.ToString() + ws.Cells["K" + (k + 1).ToString()].Value.ToString();//K列是中古韻母

                if (ws.Cells["G" + (k + 1).ToString()].Value != null)
                {
                    string shangguin = ws.Cells["G" + (k + 1).ToString()].Value.ToString();//G列是上古擬音
                    
                    if ( zhongguüinmu == "二麥")
                    {
                        if (shangguin.Contains("ak"))
                        {                          
                            ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ak", "ɛk");
                        }
                        
                    }
                    else if (zhongguüinmu == "三陌" || zhongguüinmu == "二陌")
                    {
                        if (shangguin.Contains("ɛk"))
                        {
                            ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ɛk", "ak");
                        }
                    }
                }
            }
            wb.Save(@"D:\shangguliin.xlsx");

        }

        private static void Si5Mä5()
        {
            Workbook wb = new Workbook(@"D:\《廣韻》形聲考李.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9912, 1);
            int ii;
            for (int k = 2; k < 9912; k++)
            {
                string zhongguüinmu = ws.Cells["I" + (k + 1).ToString()].Value.ToString() + ws.Cells["K" + (k + 1).ToString()].Value.ToString();//K列是中古韻母

                if (ws.Cells["G" + (k + 1).ToString()].Value != null)
                {
                    string shangguin = ws.Cells["G" + (k + 1).ToString()].Value.ToString();//G列是上古擬音

                    if (zhongguüinmu == "三藥")
                    {
                        if (shangguin.Contains("ak"))
                        {
                            ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace("ak", "øk");
                        }
                    }                   
                }
            }
            wb.Save(@"D:\shangguliin.xlsx");

        }
    }
}


