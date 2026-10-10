// See https://aka.ms/new-console-template for more information
using System.Collections.Generic;
using System.Text;

Console.WriteLine("Hello, World!");
const string uen2jän4ja5 = @"D:\MyDocument\音韻學\st sk\探索圓脣無介音W\liänmiän\";
Console.OutputEncoding = Encoding.UTF8;

string path = uen2jän4ja5 + "dazydiän.txt";
const string zy4biao3 = uen2jän4ja5 + "langjin.dict.yaml";
if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
{
    Console.Error.WriteLine("Pfad ungültig oder Datei existiert nicht.");
    return 1;
}

try
{
    var results = ExtractBeforeChar(path, '】', Encoding.UTF8);
    List<string[]> zy4zhu4List = new List<string[]>(); ;
    Console.WriteLine($"Gefundene Zeilen: {results.Count}");
    foreach (var item in results)
    {//獲得所有帶標注的字，如“昨 zo5 文”
        var line = item.Replace("【", " "); // Entfernt Leerzeichen)
        var zyinzhu=line.Split(' ');
        string zy = zyinzhu[0].Replace(")", "").Replace("(", "");
       
        var itiyz=Uen3Jwn4Li3Sou1Zy4(zy.ToList());
        foreach (var it in itiyz)
        {
            line = it + " " + zyinzhu[1]+ " " + zyinzhu[2];
            zy4zhu4List.Add(new string[3] {it, zyinzhu[1], zyinzhu[2] });
            Console.WriteLine(line);
        }        
    }

    ModifyFile(zy4biao3, zy4zhu4List, Encoding.UTF8);


    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Fehler beim Lesen der Datei: {ex.Message}");
    return 2;
}

static void ModifyFile(string filePath, List<string[]> zy1zhu4List, Encoding encoding)
{
    string backupPath = filePath + ".bak";
    File.Copy(filePath, backupPath, overwrite: true);

    var hang = File.ReadAllLines(filePath, encoding);
    for (int i = 0; i < hang.Length; i++)
    {
        foreach (var zy1zhu4 in zy1zhu4List)
        {
            var zi4ho2in1=hang[i].Split('\t');
            if (zi4ho2in1[0].Trim()==zy1zhu4[0].Trim() && zi4ho2in1[1].Trim() == zy1zhu4[1].Trim())
            {
                hang[i] += "\t" + zy1zhu4[2].Replace("文","書");
                break; // Wenn eine Übereinstimmung gefunden wurde, brechen Sie die Schleife ab, um doppelte Anhänge zu vermeiden.
            }
        }
    }
    File.WriteAllLines(filePath, hang, encoding);
}

// Liest zeilenweise, findet 'targetChar' und gibt die Teilzeichenfolge vor dem Zeichen (ohne das Zeichen) zurück.
static List<string> ExtractBeforeChar(string filePath, char targetChar, Encoding encoding)
{
    var list = new List<string>();
    using var sr = new StreamReader(filePath, encoding);
    string? line;
    while ((line = sr.ReadLine()) != null)
    {
        int idx = line.IndexOf(targetChar);
        if (idx >= 0)
        {
            string part = line.Substring(0, idx).Trim();
            if (!string.IsNullOrEmpty(part))
                list.Add(part);
        }
    }
    return list;
}

List<string> Uen3Jwn4Li3Sou1Zy4(List<char> zy)//漢字串中列出漢字
{
    List<string> res = new List<string>();
    List<string> bu5chong2di5zy4 = new List<string>();
    for (int i = 0; i < zy.Count; i++)
    {
        var c = zy[i];
        string hangzy = c.ToString();
        if (((short)c) > -20000 && ((short)c) < 0)
        {//拓展字
            hangzy = c.ToString() + zy[i + 1];
            i++;
        }
        if (!"(=)12".Contains(c) && !bu5chong2di5zy4.Contains(hangzy) && !res.Contains(hangzy))
        {
            res.Add(hangzy.ToString());           
        }
        bu5chong2di5zy4.Add(hangzy);
    }
    Console.WriteLine();
    return res;
}
