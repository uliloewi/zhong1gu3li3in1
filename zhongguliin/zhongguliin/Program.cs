using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace zhongguliin
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
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
                    {"章","cç" }, {"昌","cçʰ" }, {"常","ɟʝ" }, {"書","ç" }, {"船","ʝ" },
                    {"以","ʎ" }, {"日","ɲ" },
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
                    {"祭","ɛʎ" },//重紐皆三等
                    {"泰","aʎ" },//皆一等
                    {"佳","ɛ" },//皆二等 
                    {"皆","æi"},//皆二等
                    {"夬","aʎ"},//皆二等
                    {"咍","ɒi" },//皆一開
                    {"灰","ɒi" },//皆一合
                    {"廢","ɒʎ" },//皆三等
                    // 臻攝 
                    {"眞","ɨn" },{"眞入","ɨt" },//重紐皆三等，三A改i韻腹
                    {"臻","ən" },{"臻入","ət" },//皆三開
                    {"欣","on" },{"欣入","ot" },//皆三開
                    {"文","ən" },{"文入","ət" },//皆三合
                    {"痕","on" },{"痕入","ot" },//皆一開
                    {"魂","on" },{"魂入","ot" },//皆合開
                    {"諄","ɨn" },{"諄入","ɨt" },//皆三合，即眞B合
                    // 山攝 
                    {"寒","an" },{"寒入","at" },//皆一開，除了䔾三開
                    {"桓","an" },{"桓入","at" },//皆一開
                    {"刪","an" },{"刪入","at" },//皆二等
                    {"山","æn" },{"山入","æt" },//皆二等
                    {"元","ɒn" },{"元入","ɒt" },//皆三等
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
                    {'c',"t" },{'ɟ',"d" },{'ç',"sj" },{'ʝ',"zj" },{'ʎ',"j" },{'ɲ',"nj" },
                    {'u',"u" },{'ʅ',"r" },{'ʯ',"w" },{'ɨ',"y" },{'ʉ',"v" },{'i',"i" },{'y',"ü" },
                    {'o',"o" },{'ɪ',"ï" },{'ə',"ë" },{'ɛ',"e" },{'a',"a" },{'æ',"ä" },{'ɒ',"ö" },
                };

                Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(@"D:\Downloads\Guangyun_Langjin_pulish_Alphabetic.2.0.xlsx");
                Worksheet ws = wk.Worksheets[0];

                var dt = ws.Cells.ExportDataTable(0, 0, 4000, 9);

                Aspose.Cells.Workbook wkwrite = new Aspose.Cells.Workbook();
                Worksheet wswrite = wkwrite.Worksheets[0];

                for (int k = 0; k < dt.Rows.Count && dt.Rows[k][2].ToString() != ""; k++)
                {
                    if (dt.Rows[k][3].ToString().Contains("等"))
                        continue;
                    string shengmu = sheng[dt.Rows[k][2].ToString().Substring(0, 1)];
                    string üinshou = jäin[(dt.Rows[k][5].ToString().Contains("A") || dt.Rows[k][5].ToString().Contains("幽") || dt.Rows[k][5].ToString().Contains("清") ? "四" : dt.Rows[k][3].ToString().Substring(0, 1)) + dt.Rows[k][4].ToString().Substring(0, 1)];
                    string üinmu = dt.Rows[k][6].ToString().Contains("入") ? üin[dt.Rows[k][5].ToString().Substring(0, 1) + dt.Rows[k][6].ToString().Substring(0, 1)] : üin[dt.Rows[k][5].ToString().Substring(0, 1)];
                    string diaozhr = diao[dt.Rows[k][6].ToString()];
                    string inqüin = shengmu + üinshou + üinmu;
                    string inbiao = inqüin.Replace("ii", "i")
                        .Replace("uu", "u").Replace("yy", "y")
                        .Replace("ɨɨ", "ɨ")
                        .Replace("ʉʉ", "ʉ")
                        .Replace("ʅʅ", "ʅ")
                        .Replace("ʯʯ", "ʯ")
                        .Replace("iɨ", "i");//改眞侵重紐韻腹
                    wswrite.Cells.Rows[k][0].Value = inbiao + diaozhr;
                    string pinin = "";
                    foreach (char c in inbiao)
                        pinin += inbiao2pinin[c];

                    Console.WriteLine(k);
                    if (dt.Rows[k][6].ToString().Contains("上") || dt.Rows[k][6].ToString().Contains("去"))
                    {
                        BiaoShenDiao(ref pinin, dt.Rows[k][6].ToString());      
                    }
                    wswrite.Cells.Rows[k][1].Value = pinin;
                }
                wkwrite.Save("d:\\d2.xls");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        
        private static void BiaoShenDiao(ref string pinin, string shendiao)//標聲調
        {
            string üinfu = "aeiouyäöëï";
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

    }
}


