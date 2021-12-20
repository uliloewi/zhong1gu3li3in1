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
            Dictionary<string, string> sheng = new Dictionary<string, string>()
            {
                {"幫","p" }, {"滂","pʰ" }, {"並","b" }, {"明","m" },
                {"端","t" }, {"透","tʰ" }, {"定","d" }, {"泥","n" },
                {"精","ts" }, {"清","tsʰ" }, {"從","dz" }, {"心","s" }, {"邪","z" },
                {"見","k" }, {"溪","kʰ" }, {"羣","g" }, {"疑","ŋ" },
                {"影","ʔ" },
                {"曉","x" }, {"匣","ɣ" },
                {"云","" },
                {"來","l" },
                 //翹舌映二等
                {"知","ʈ" }, {"徹","ʈʰ" }, {"澄","ɖ" }, {"孃","ɳ" },
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
                {"平","˧˧" }, {"上","˨˦" }, {"去","˧˩" }, {"入","˥" },
            };
            Dictionary<string, string> üin = new Dictionary<string, string>()
            {
                {"東","uŋ" },//皆一三開
                {"冬","oŋ" },//皆一開
                {"鍾","oŋ" },//皆三開
                {"江","ɒŋ" },//皆二開
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
                {"祭","ɛɿ" },//重紐皆三等
                {"泰","æɿ" },//皆一等
                {"佳","ɛ" },//皆二等 
                {"皆","ai"},//皆二等
                {"夬","æɿ"},//皆二等
                {"咍","ai" },//皆一開
                {"灰","ai" },//皆一合
                {"廢","æɿ" },//皆三等
                // 臻攝 
                {"眞","ɨn" },//重紐皆三等
                {"臻","ən" },//皆三開
                {"欣","on" },//皆三開
                {"文","ən" },//皆三合
                {"痕","on" },//皆一開
                {"魂","on" },//皆合開
                {"諄","ʉn" },//皆三合，即眞B合
                // 山攝 
                {"寒","an" },//皆一開，除了䔾三開
                {"桓","an" },//皆一開
                {"刪","an" },//皆二等
                {"山","æn" },//皆二等
                {"元","an" },//皆三等
                {"仙","æn" },//重紐皆三等
                {"先","ɛn" },//皆四等
                // 效攝 
                {"蕭","æu" },//皆四開
                {"宵","au" },//重紐皆三開
                {"肴","au" },//皆二開
                {"豪","au" },//皆一開
                // 果攝 
                {"歌","ɒ" },//皆一開
                {"戈","ɒ" },//皆一三等
                // 假攝 
                {"麻","a" },//皆一等或三開
                // 宕攝 
                {"陽","aŋ" },//皆三等
                {"唐","aŋ" },//皆一等
                // 梗攝 
                {"庚","æŋ" },//皆二三等
                {"耕","ɛŋ" },//皆二等
                {"清","æŋ" },//皆三A
                {"青","ɛŋ" },//皆四等
                // 曾攝 
                {"蒸","ɨŋ" },//皆三等
                {"登","əŋ" },//皆一等
                // 流攝 
                {"尤","u" },//皆三開
                {"侯","u" },//皆一開
                {"幽","iu" },//皆三A開
                // 深攝 
                {"侵", "əm" },//重紐皆三開
                // 咸攝 
                {"覃","ɒm" },//皆一開
                {"談","am" },//皆一開
                {"咸","æm" },//皆二開
                {"銜","am" },//皆二開
                {"凡","am" },//皆三合
                {"鹽","æm" },//重紐皆三開
                {"嚴","am" },//皆三開
                {"添","am" },//皆四開
            };


            Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(@"D:\projects\dotnet\test\Book1.xlsx");
            Worksheet ws = wk.Worksheets[0];

            var dt = ws.Cells.ExportDataTable(0, 0, 9, 9);

            for (int k = 0; k < dt.Rows.Count && dt.Rows[k][2].ToString() != ""; k++)
            {

                string shengmu = sheng[dt.Rows[k][2].ToString().Substring(0,1)];
                string üinshou = jäin[(dt.Rows[k][5].ToString().Contains("A") || dt.Rows[k][5].ToString().Contains("幽") || dt.Rows[k][5].ToString().Contains("清") ? "四" : dt.Rows[k][3].ToString().Substring(0,1)) + dt.Rows[k][4].ToString().Substring(0,1)];
                string üinmu = üin[dt.Rows[k][5].ToString().Substring(0, 1)];
                string shengdiao = diao[dt.Rows[k][6].ToString()];
                string inqüin = shengmu + üinshou + üinmu + shengdiao;
                string inzie = inqüin.Replace("ii", "i").Replace("uu", "u").Replace("yy", "y").Replace("ɨɨ", "ɨ").Replace("ʉʉ", "ʉ").Replace("ʅʅ", "ʅ").Replace("ʯʯ", "ʯ");
                   
                Console.WriteLine(inzie);
            }

            //using (StreamReader sr = new StreamReader(@"D:\projects\dotnet\test\Unbenannt 1.csv", Encoding.Default, true))
            //{
            //    string currentLine;
            //    // currentLine will be null when the StreamReader reaches the end of file
            //    while ((currentLine = sr.ReadLine()) != null)
            //    {
            //        // Search, case insensitive, if the currentLine contains the searched keyword
            //        if (currentLine.IndexOf("I/RPTGEN", StringComparison.CurrentCultureIgnoreCase) >= 0)
            //        {
            //            Console.WriteLine(currentLine);
            //        }
            //    }
            //}

            Console.WriteLine("Hello World!");
        }
    }
}


