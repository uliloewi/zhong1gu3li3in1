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
            try
            {
                Console.TreatControlCAsInput = true;
                Console.WriteLine("生成擬音請按'1'；生成字表請按'2'；生成多音字統計請按'3':");
                ConsoleKeyInfo cki = Console.ReadKey();
                var str = cki.Key.ToString();
                Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(@"D:\\Guangyun_Langjin_Zhonggu.1.0.xlsx");
                Worksheet ws = wk.Worksheets[0];
                var dt = ws.Cells.ExportDataTable(0, 0, 4000, 19);
                
                switch (str.Substring(str.Length - 1))
                {
                    case "1":
                        Dictionary<string, string> sheng = new Dictionary<string, string>()
                        {
                            {"幫","p" }, {"滂","pʰ" }, {"並","b" }, {"明","m" },
                            {"端","t" }, {"透","tʰ" }, {"定","d" }, {"泥","n" },
                            {"精","ts" }, {"清","tsʰ" }, {"從","dz" }, {"心","s" }, {"邪","z" },
                            {"見","k" }, {"溪","kʰ" }, {"群","g" }, {"羣","g" }, {"疑","ŋ" },
                            {"影","ʔ" },
                            {"曉","x" }, {"匣","ɣ" },
                            {"云","" },
                            {"來","l" },
                             //翹舌映二等
                            {"知","ʈ" }, {"徹","ʈʰ" }, {"澄","ɖ" }, {"孃","ɳ" },{"娘","ɳ" },
                            {"莊","tʂ" }, {"初","tʂʰ" }, {"崇","dʐ" }, {"生","ʂ" }, {"俟","ʐ" },
                             //硬腭映三等
                            {"章","tɕ" }, {"昌","tɕʰ" }, {"常","dʑ" }, {"書","ɕ" }, {"船","ʑ" },
                            {"以","j" }, {"日","ɲ" },
                        };
                        Dictionary<string, string> jäin = new Dictionary<string, string>()
                        {
                            {"一開","" }, {"一合","u" }, {"二開","ʅ" }, {"二合","ʯ" },
                            {"三開","ɨ" }, {"三合","ʉ" }, {"三開A","i" }, {"三合A","y" }, {"四開","i" }, {"四合","y" },
                        };
                        Dictionary<string, string> diao = new Dictionary<string, string>()
                        {
                            //{"平","" }, {"上","x" }, {"去","h" }, {"入","" },
                            {"平","˧˧" }, {"上","˨˦ 'x" }, {"去","˧˩ 'h" }, {"入","˥" },
                        };
                        Dictionary<string, string> üin = new Dictionary<string, string>()
                        {
                            {"東","uŋ" },{"東入","uk" },//皆一三開
                            {"冬","oŋ" },{"冬入","ok" },//皆一開
                            {"鍾","oŋ" },{"鍾入","ok" },//皆三開
                            {"江","ɒŋ" },{"江入","ɒk" },//皆二開
                            {"支","ɪ" },//重紐皆三等
                            {"脂","i" },//重紐皆三等
                            {"之","ɨ" },//皆三開
                            {"微","əi" },//皆三等
                            // 遇攝 
                            {"魚","o" },//皆三開
                            {"虞","o" },//皆三合
                            {"模","o" },//皆一開
                            // 蟹攝
                            {"齊","ɛi" },//皆四等
                            {"祭","ɛj" },//重紐皆三等
                            {"泰","aj" },//皆一等
                            {"佳","ɛ" },//皆二等 
                            {"皆","æi"},//皆二等
                            {"夬","aj"},//皆二等
                            {"咍","ɒi" },//皆一開
                            {"灰","ɒi" },//皆一合
                            {"廢","ɒj" },//皆三等
                            // 臻攝 
                            {"眞","ɨn" },{"眞入","ɨt" },//重紐皆三等，三A改i韻腹
                            {"臻","ɨn" },{"臻入","ɨt" },//皆三開
                            {"欣","ən" },{"欣入","ət" },//皆三開
                            {"文","ən" },{"文入","ət" },//皆三合
                            {"痕","on" },{"痕入","ot" },//皆一開
                            {"魂","on" },{"魂入","ot" },//皆合開
                            {"諄","ɨn" },{"諄入","ɨt" },//皆三合，三A改i韻腹
                            // 山攝 
                            {"寒","an" },{"寒入","at" },//皆一開，除了䔾三開
                            {"桓","an" },{"桓入","at" },//皆一開
                            {"刪","an" },{"刪入","at" },//皆二等
                            {"山","æn" },{"山入","æt" },//皆二等
                            {"元","on" },{"元入","ot" },//皆三等
                            {"仙","æn" },{"仙入","æt" },//重紐皆三等
                            {"先","ɛn" },{"先入","ɛt" },//皆四等
                            // 效攝 
                            {"蕭","ɛu" },//皆四開
                            {"宵","æu" },//重紐皆三開
                            {"肴","au" },//皆二開
                            {"豪","au" },//皆一開
                            // 果攝 
                            {"歌","ɒ" },//皆一開
                            {"戈","ɒ" },//皆一三等
                            // 假攝 
                            {"麻","a" },//皆一等或三開
                            // 宕攝 
                            {"陽","aŋ" },{"陽入","ak" },//皆三等
                            {"唐","aŋ" },{"唐入","ak" },//皆一等
                            // 梗攝 
                            {"庚","æŋ" },{"庚入","æk" },//皆二三等
                            {"耕","ɛŋ" },{"耕入","ɛk" },//皆二等
                            {"清","æŋ" },{"清入","æk" },//皆三A
                            {"青","ɛŋ" },{"青入","ɛk" },//皆四等
                            // 曾攝 
                            {"蒸","ɨŋ" },{"蒸入","ɨk" },//皆三等
                            {"登","əŋ" },{"登入","ək" },//皆一等
                            // 流攝 
                            {"尤","u" },//皆三開
                            {"侯","u" },//皆一開
                            {"幽","iu" },//皆三A開
                            // 深攝 
                            {"侵","ɨm" },{"侵入","ɨp" },//重紐皆三開，三A改i韻腹
                            // 咸攝 
                            {"覃","ɒm" },{"覃入","ɒp" },//皆一開
                            {"談","am" },{"談入","ap" },//皆一開
                            {"咸","æm" },{"咸入","æp" },//皆二開
                            {"銜","am" },{"銜入","ap" },//皆二開
                            {"凡","ɒm" },{"凡入","ɒp" },//皆三合
                            {"鹽","æm" },{"鹽入","æp" },//重紐皆三開
                            {"嚴","ɒm" },{"嚴入","ɒp" },//皆三開
                            {"添","ɛm" },{"添入","ɛp" },//皆四開
                        };

                        Dictionary<char, string> inbiao2pinin = new Dictionary<char, string>()
                        {
                            {'ʰ',"h" },{'p',"p" },{'b',"b" },{'m',"m" },
                            {'t',"t" },{'d',"d" },{'n',"n" },{'s',"s" },{'z',"z" },{'l',"l" },
                            {'k',"k" },{'g',"g" },{'ŋ',"ng" },{'ʔ',"q" },{'x',"h" },{'ɣ',"x" },
                            {'ʈ',"tc" },{'ɖ',"dc" },{'ɳ',"n" },{'ʂ',"sc" },{'ʐ',"zc" },
                            {'c',"t" },{'ɟ',"d" },{'ɕ',"sj" },{'ʑ',"zj" },{'j',"j" },{'ɲ',"nj" },
                            {'u',"u" },{'ʅ',"r" },{'ʯ',"w" },{'ɨ',"y" },{'ʉ',"ü" },{'i',"i" },{'y',"v" },
                            {'o',"o" },{'ɪ',"ï" },{'ə',"ë" },{'ɛ',"e" },{'a',"a" },{'æ',"ä" },{'ɒ',"ö" },
                        };                        

                        Aspose.Cells.Workbook wkwrite = new Aspose.Cells.Workbook();
                        Worksheet wswrite = wkwrite.Worksheets[0];
                        List<String> ls = new List<string>();
                        for (int k = 0; k < dt.Rows.Count && dt.Rows[k][2].ToString() != ""; k++)
                        {
                            if (dt.Rows[k][3].ToString().Contains("等"))
                                continue;

                            if (!ls.Contains(dt.Rows[k][11].ToString()))
                            {
                                ls.Add(dt.Rows[k][11].ToString());
                            }

                            string change = "";
                            string newval = "";
                            if (dt.Rows[k][13].ToString().StartsWith("改"))
                            {
                                var val = dt.Rows[k][13].ToString().Split(" ");
                                change = val[0].Substring(val[0].Length - 1, 1);
                                switch (change)
                                {
                                    case "口":
                                    case "母":
                                    case "聲":
                                        newval = val[0].Substring(1, 1);
                                        break;
                                    case "韻":
                                        newval = val[0].Substring(1, val[0].Length - 2);

                                        break;
                                }

                            }

                            string shengmu = sheng[dt.Rows[k][2].ToString().Substring(0, 1)];
                            string shengmu2 = change == "母" || newval.Contains("母") ? sheng[newval.Substring(0, 1)] : shengmu;
                            string üinshou = jäin[(dt.Rows[k][5].ToString().Contains("A") || dt.Rows[k][5].ToString().Contains("幽") || dt.Rows[k][5].ToString().Contains("清") ? "四" : dt.Rows[k][3].ToString().Substring(0, 1)) + dt.Rows[k][4].ToString().Substring(0, 1)];
                            
                            string üinshou2 = üinshou;
                            if (change != "聲" )
                            { 
                                if (change == "口")
                                üinshou2 = jäin[(dt.Rows[k][5].ToString().Contains("A") || dt.Rows[k][5].ToString().Contains("幽") || dt.Rows[k][5].ToString().Contains("清") ? "四" : dt.Rows[k][3].ToString().Substring(0, 1)) + newval];                  

                                else if (newval.Contains("A") || newval.Contains("B") || newval.Contains("幽"))
                                    üinshou2 = jäin[(newval.Contains("A") || newval.Contains("幽") ? "四" : dt.Rows[k][3].ToString().Substring(0, 1)) + dt.Rows[k][4].ToString().Substring(0, 1)];
                                else if (newval.Contains("等"))
                                    üinshou2 = jäin[newval.Substring(newval.IndexOf("等") - 1,1) + dt.Rows[k][4].ToString().Substring(0, 1)];
                            }
                            string üinmu = dt.Rows[k][6].ToString().Contains("入") ? üin[dt.Rows[k][5].ToString().Substring(0, 1) + dt.Rows[k][6].ToString().Substring(0, 1)] : üin[dt.Rows[k][5].ToString().Substring(0, 1)];
                            string üinmu2 = üinmu;
                            if (newval.Length > 0)
                            { 
                                string sinüin = newval.Contains("A") || newval.Contains("B") ? newval.Substring(newval.Length - 2, 1) : newval.Substring(newval.Length - 1, 1);
                                üinmu2 = change == "韻" ? (dt.Rows[k][6].ToString().Contains("入") ? üin[sinüin + dt.Rows[k][6].ToString().Substring(0, 1)] : üin[sinüin]) : üinmu;
                            }
                            string diaozhr = diao[dt.Rows[k][6].ToString()];
                            string diaozhr2 = change != "聲" ? diaozhr : diao[newval];
                            string inqüin = shengmu + üinshou + üinmu;
                            string inqüin2 = shengmu2 + üinshou2 + üinmu2;
                            string inbiao = inqüin.Replace("ii", "i")
                                .Replace("uu", "u").Replace("yy", "y")
                                .Replace("ɨɨ", "ɨ").Replace("ʉʉ", "ʉ")
                                .Replace("ʅʅ", "ʅ").Replace("ʯʯ", "ʯ")
                                .Replace("iɨ", "i")//改眞諄侵重四開韻腹
                                .Replace("yɨ", "yi");//改眞諄侵重四合韻腹
                            string inbiao2 = inqüin2.Replace("ii", "i")
                                .Replace("uu", "u").Replace("yy", "y")
                                .Replace("ɨɨ", "ɨ").Replace("ʉʉ", "ʉ")
                                .Replace("ʅʅ", "ʅ").Replace("ʯʯ", "ʯ")
                                .Replace("iɨ", "i")//改眞諄侵重四開韻腹
                                .Replace("yɨ", "yi");//改眞諄侵重四合韻腹
                            wswrite.Cells.Rows[k][0].Value = inbiao + diaozhr;
                            wswrite.Cells.Rows[k][2].Value = inbiao2 + diaozhr2;
                            string pinin = "";
                            string pinin2 = "";
                            foreach (char c in inbiao)
                                pinin += inbiao2pinin[c];
                            foreach (char c in inbiao2)
                                pinin2 += inbiao2pinin[c];
                            var inzie2 = pinin2;

                            Console.WriteLine(k);
                            if (dt.Rows[k][6].ToString().Contains("上") || dt.Rows[k][6].ToString().Contains("去"))
                            {
                                BiaoShenDiao(ref pinin, dt.Rows[k][6].ToString());
                                BiaoShenDiao(ref pinin2, dt.Rows[k][6].ToString());
                            }
                            if (change == "聲")
                            {
                                pinin2 = inzie2;
                                if (newval.Contains("上") || newval.Contains("去"))
                                {
                                    BiaoShenDiao(ref pinin2, newval);
                                }
                            }

                            wswrite.Cells.Rows[k][1].Value = pinin;
                            wswrite.Cells.Rows[k][3].Value = pinin2;
                        }
                        Console.WriteLine(ls.Count);
                        wkwrite.Save("d:\\d2.xls");
                        break;
                    case "2":
                        List<string> lineList = new List<string>();
                        for (int k = 0; k < dt.Rows.Count && dt.Rows[k][2].ToString() != ""; k++)
                        {
                            if (dt.Rows[k][3].ToString().Contains("等"))
                                continue;
                            int checkSurrogate = 1;
                            char firstSurrogate = ' ';
                            foreach (char ch in dt.Rows[k][12].ToString())
                            {
                                string pinin = dt.Rows[k][10].ToString();
                                if (ch == '(' || ch == ')')
                                {
                                    continue;
                                }
                                if (char.IsHighSurrogate(ch) && checkSurrogate == 1)
                                {//複雜字的前半個
                                    firstSurrogate = ch;
                                    checkSurrogate++;
                                }
                                else if (checkSurrogate == 2)
                                {//複雜字的後半個
                                    string complexChar = String.Concat(firstSurrogate, ch);
                                    lineList.Add(complexChar + "	" + dt.Rows[k][10].ToString());
                                    checkSurrogate = 1;
                                }
                                else
                                    lineList.Add(ch + "	" + dt.Rows[k][10].ToString());
                            }
                        }
                        File.WriteAllLines("d:\\tongguang.yaml", lineList) ;
                        break;
                    case "3":
                        List<string> dict = new List<string>();
                        for (int k = 0; k < dt.Rows.Count && dt.Rows[k][2].ToString() != ""; k++)
                        {
                            if (dt.Rows[k][3].ToString().Contains("等"))
                                continue;
                            int checkSurrogate = 1;
                            char firstSurrogate = ' ';
                            foreach (char ch in dt.Rows[k][12].ToString())
                            {
                                
                                string pinin = dt.Rows[k][10].ToString();
                                if (ch == '(' || ch == ')'|| ch == '+' || ch == '*')
                                {
                                    continue;
                                }
                                if (char.IsHighSurrogate(ch) && checkSurrogate == 1)
                                {//複雜字的前半個
                                    firstSurrogate = ch;
                                    checkSurrogate++;
                                }
                                else if (checkSurrogate == 2)
                                {//複雜字的後半個
                                    string complexChar = String.Concat(firstSurrogate, ch);
                                    if (!dict.Any(x => x.StartsWith(complexChar)))
                                        dict.Add(complexChar + "," + dt.Rows[k][10].ToString());
                                    else
                                    {
                                        string s = dict.Where(x => x.StartsWith(complexChar)).First() + "," + dt.Rows[k][10].ToString();
                                        dict.RemoveAll(x => x.StartsWith(complexChar));
                                        dict.Add(s);
                                    }
                                    checkSurrogate = 1;
                                }
                                else
                                {
                                    if (!dict.Any(x => x.StartsWith(ch)))
                                        dict.Add(ch + "," + dt.Rows[k][10].ToString());
                                    else
                                    {
                                        string s = dict.Where(x => x.StartsWith(ch)).First() + "," + dt.Rows[k][10].ToString();
                                        dict.RemoveAll(x => x.StartsWith(ch));
                                        dict.Add(s);
                                    }
                                }
                            }
                        }
                        dict.RemoveAll(x => x.Count(f => f == ',') == 1);
                        File.WriteAllLines("d:\\do1in1zy4.csv", dict.OrderBy(x => x.Count(f => f == ',')));

                        break;
                    default:
                        break;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        
        private static void BiaoShenDiao(ref string pinin, string shendiao)//標聲調
        {
            string üinfu = "aeiouyäöüëï";
            int idx = 10;
            foreach (char ch in üinfu)
            {
                int loc = pinin.IndexOf(ch);
                if (loc > -1 && loc < idx)
                    idx = loc;
            }
            if (idx + 1 < pinin.Length && üinfu.Contains(pinin.Substring(idx + 1, 1)))
                idx++;
            if (pinin.Contains('a')) idx = pinin.IndexOf('a');
            else if (pinin.Contains('ä')) idx = pinin.IndexOf('ä');
            else if (pinin.Contains('ö')) idx = pinin.IndexOf('ö');
            else if (pinin.Contains('e')) idx = pinin.IndexOf('e');
            else if (pinin.Contains('ë')) idx = pinin.IndexOf('ë');
            else if (pinin.Contains('u')) idx = pinin.IndexOf('u');
            string zhuüänin = pinin.Substring(idx, 1);
            string sinzyfu = zhuüänin + (shendiao.Contains("上") ? "\u0301" : "\u0300");
            sinzyfu.Normalize();
            pinin = pinin.Replace(zhuüänin, sinzyfu);
        }

        private static List<string> ChangeStringList(List<string> dict, string ch, string du5in1)
        {
            List<string> res = new List<string>(dict);
            if (!res.Any(x => x.StartsWith(ch)))
                res.Add(ch + "," + du5in1);
            else
            {
                string s = res.Where(x => x.StartsWith(ch)).First() + "," + du5in1;
                res.RemoveAll(x => x.StartsWith(ch));
                res.Add(s);
            }
            return res;
        }
    }
}


