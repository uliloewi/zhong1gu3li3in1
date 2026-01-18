// See https://aka.ms/new-console-template for more information
using System.Reflection;
using System.Text;
using Aspose.Cells;
using Microsoft.Office.Interop.Word;

Console.WriteLine("Hello, World!");
const string uen2jän4ja5 = @"D:\MyDocument\音韻學\st sk\探索圓脣無介音W\";
Console.OutputEncoding = Encoding.UTF8;
Workbook wk = new Workbook(uen2jän4ja5 + "廣韻字上古音形考.xlsx");
Worksheet ws = wk.Worksheets[0];
//CheckDen(ws);
//int length = CheckDoubleMapping(ws);

Application wordApp = new Application();        

Document wordDoc = wordApp.Documents.Open(uen2jän4ja5+ "a.docx");

// Perform operations with the document here (e.g., read text)
Console.WriteLine("Document opened: " + wordDoc.Name);

WriteFromExcelToWord(ws,0,wordDoc);
//Add(wordDoc);

wordDoc.Save();
// 3. Close the document and save changes if necessary
wordDoc.Close();

static int CheckDoubleMapping(Worksheet ws)
{
    Dictionary<string, List<string>> Mapping = new Dictionary<string, List<string>>();
    int res = 3;
    int nullcount = 0;
    while (LoopCondition(ws, res)) //"G"列是上古音
    {
        if (ws.Cells["G" + res.ToString()].Value != null)//"G"列是上古音
        {
            string k = ws.Cells["G" + res.ToString()].Value.ToString();//"G"列是上古音                    
            if (ws.Cells["N" + res.ToString()].Value != null)//"N"列是中古音
            {
                string v = ws.Cells["N" + res.ToString()].Value.ToString() + ws.Cells["K" + res.ToString()].Value.ToString();//"K"列是中古韻
                if (!Mapping.ContainsKey(k))
                {
                    Mapping.Add(k, new List<string>());
                }
                if (!Mapping[k].Contains(v))
                {
                    Mapping[k].Add(v);
                }
                nullcount = 0;
            }
        }
        else
        {
            nullcount++;
        }
        if (nullcount > 4)
            break;
        res++;
    }
    foreach (var kv in Mapping.Where(kv => kv.Key.Trim() != ""))
    {
        if (kv.Value.Count > 1)
        {
            Console.WriteLine("以下的上古擬音對映多個中古音，請重擬！");
            Console.Write(kv.Key + " ");
            foreach (var v in kv.Value)
                Console.Write(v + " ");
            Console.WriteLine();
        }
    }
    return res - nullcount;

}

static bool LoopCondition(Worksheet ws, int res)
{
    return ws.Cells["G" + res.ToString()].Value == null ||
        !String.IsNullOrWhiteSpace(ws.Cells["G" + res.ToString()].Value.ToString()) ||//"G"列是上古音
        (ws.Cells["P" + res.ToString()].Value != null &&
        !String.IsNullOrWhiteSpace(ws.Cells["P" + res.ToString()].Value.ToString()));
}


static void RemoveRow(Document doc, int n)
{

    Table currentTable = doc.Tables[1];
    // Get total rows
    int totalRows = currentTable.Rows.Count;

    // Loop backwards from the last row until n rows have been deleted
    // Ensure you don't try to delete more rows than exist
    for (int i = totalRows; i > (totalRows - n) && i > 0; i--)
    {
        currentTable.Rows[i].Delete();
    }
}

static void AddParagraph(Document doc, string text, int bold, int spaceBefore)
{
    Paragraph para = doc.Paragraphs.Add();
    // 3. Set the text for the paragraph
    para.Range.Text = text;
    // 2. Set the distance above (in points)
    para.Format.SpaceBefore = spaceBefore;

    // 4. (Optional) Apply formatting
    para.Range.Font.Bold = bold; // 1 is true, 0 is false
    para.Format.SpaceAfter = 12; // Adds 12 points of space after the paragraph

    // 5. To ensure subsequent text starts on a NEW paragraph:
    para.Range.InsertParagraphAfter();

}

static void WriteFromExcelToWord(Worksheet ws, int colIndex, Document doc)
{
    int cnt = 1;
    for (int row = 2; row <= ws.Cells.MaxDataRow; row++)
    {
        var cell = ws.Cells[row, colIndex];
        if (cell == null || cell.Value == null) return;
        string shenpang = cell.Value.ToString().Replace("|", "”“");//如：聲旁“丿”“乀”
        string paraTitle = cnt.ToString() + ". 聲旁“" + shenpang + "”";
        AddParagraph(doc, paraTitle, 1, 40);//如：2. 聲旁“丿”“乀”
        Console.WriteLine(paraTitle);
        cnt++;
        AddParagraph(doc, "這族形聲字在上古的聲母", 0, 10);
        Table currentTable = CopyTable(doc);
        if (cell.IsMerged)
        {
            var range = cell.GetMergedRange();
            //var originalValue = cell.Value;
            //ws.Cells.UnMerge(range.FirstRow, range.FirstColumn, range.RowCount, range.ColumnCount);
            for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
            {
                CopyCellValues(ws, currentTable, r == range.FirstRow, r);
            }
            row = range.FirstRow + range.RowCount - 1;
        }
        else //單行如“𠨲”
        {
            CopyCellValues(ws, currentTable, true, row);
        }

        AddParagraph(doc, "上表有0個字值得細説：", 0, 10);
    }

   
    
}

static Table CopyTable(Document doc)
{
    int tableCount = doc.Tables.Count;
    if (tableCount > 0)
    {
        Table lastTable = doc.Tables[1];

        // 2. Copy the table to the clipboard
        lastTable.Range.Copy();

        // 3. Move to the end of the document
        // We add a paragraph first so the new table isn't merged into the old one
        Microsoft.Office.Interop.Word.Range endRange = doc.Content;
        endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
        endRange.InsertParagraphAfter();

        // 4. Paste the table at the very end
        endRange.SetRange(doc.Content.End, doc.Content.End);
        endRange.Paste();
    }
    return doc.Tables[doc.Tables.Count];
}


static void CopyCellValues(Worksheet ws, Table currentTable, bool isFirstRow, int row)
{
    var lastRow = isFirstRow ? currentTable.Rows[currentTable.Rows.Count] : currentTable.Rows.Add(Missing.Value);
    CopyCell(ws, lastRow, currentTable, row, 1, 2);
    CopyCell(ws, lastRow, currentTable, row, 2, 4);
    CopyCell(ws, lastRow, currentTable, row, 3, 5);
    CopyCell(ws, lastRow, currentTable, row, 4, 6);
    CopyCell(ws, lastRow, currentTable, row, 5, 7);
    CopyCell(ws, lastRow, currentTable, row, 6, 8);
    CopyCell(ws, lastRow, currentTable, row, 7, 9);
    CopyCell(ws, lastRow, currentTable, row, 8, 10);
    CopyCell(ws, lastRow, currentTable, row, 9, 11);
    CopyCell(ws, lastRow, currentTable, row, 10, 12);
    CopyCell(ws, lastRow, currentTable, row, 11, 13);
    CopyCell(ws, lastRow, currentTable, row, 12, 14);
    CopyCell(ws, lastRow, currentTable, row, 13, 15);
    CopyCell(ws, lastRow, currentTable, row, 14, 16);
}

static void CopyCell(Worksheet ws, Microsoft.Office.Interop.Word.Row wordRow, Table currentTable, int row, int wcol, int ecol)
{
    if (ws.Cells[row, ecol].Value != null)
    {
        var cell = ws.Cells[row, ecol];
        if (cell.Value is string str && str.Length > 0)
        {
            if (wcol == 13)
                Console.WriteLine(str);
            var wholeChars = new StringBuilder(str);
            var redChars = new StringBuilder();
            var yellowChars = new StringBuilder();
            var normalChars = new StringBuilder();
            // Sammle alle roten Zeichen und entferne sie aus dem Originalstring
            for (int i = 0; i < str.Length; i++)
            {
                var c = cell.Characters(i, 1);
                if (c.Font.Color.Name == "ffff0000")//red char
                {
                    redChars.Insert(redChars.Length, str.Substring(c.StartIndex, c.Length));
                }
                else if (c.Font.Color.Name == "ffffcc00" || c.Font.Color.Name == "ffffc000")//yellow char
                {
                    yellowChars.Insert(yellowChars.Length, str.Substring(c.StartIndex, c.Length));
                }
                else
                {
                    normalChars.Insert(normalChars.Length, str.Substring(c.StartIndex, c.Length));
                }
            }

            wordRow.Cells[wcol].Range.Text = redChars.ToString() + normalChars.ToString() + yellowChars.ToString();

            if (redChars.Length > 0)
            {
                Microsoft.Office.Interop.Word.Range firstPart = wordRow.Cells[wcol].Range.Characters[1]; // Start
                firstPart.End = tryCatch( wordRow, wcol, redChars.Length);
                firstPart.Font.Italic = 0;
                firstPart.Font.Bold = 1;
                firstPart.Font.Color = WdColor.wdColorRed;
                SetFontForMainCharracters(wcol, firstPart);
            }


            int textLength = wordRow.Cells[wcol].Range.Text.Replace("\r\a", "").Length;

            if (yellowChars.Length > 0)
            {
                int start = 1;
                Microsoft.Office.Interop.Word.Range lastPart = TryCatchRange(wordRow, wcol, textLength - yellowChars.Length + 1, ref start);
                lastPart.End = tryCatch( wordRow, wcol, textLength);//wordRow.Cells[wcol].Range.Characters[textLength].End; // Last char
                lastPart.Font.Bold = 0;
                lastPart.Font.Italic = 1;
                lastPart.Font.Color = WdColor.wdColorOrange;
                SetFontForMainCharracters(wcol, lastPart);
            }

            if (textLength - redChars.Length > yellowChars.Length)
            {
                int start=1;
                Microsoft.Office.Interop.Word.Range normalPart = TryCatchRange( wordRow, wcol, redChars.Length + 1, ref start);
                if (start > -1)
                {
                    normalPart.End = tryCatch(wordRow, wcol, textLength - yellowChars.Length);
                    normalPart.Font.Bold = 0;
                    normalPart.Font.Italic = 0;
                    normalPart.Font.Color = WdColor.wdColorBlack;
                    SetFontForMainCharracters(wcol, normalPart);
                }
            }
        }        
    }
}

static int tryCatch(Microsoft.Office.Interop.Word.Row wordRow, int wcol, int length)
{

    int charaEnd=0;
    try
    {

        charaEnd = wordRow.Cells[wcol].Range.Characters[length].End;
    }
    catch
    {
        tryCatch(wordRow, wcol, length - 1);        
    }
    return charaEnd;
}

static Microsoft.Office.Interop.Word.Range TryCatchRange(Microsoft.Office.Interop.Word.Row wordRow, int wcol, int length, ref int start)
{

    Microsoft.Office.Interop.Word.Range res = wordRow.Cells[wcol].Range.Characters[start];
    start = -1;
    try
    {
        res = wordRow.Cells[wcol].Range.Characters[length];
        start = length;
    }
    catch
    {
        res = TryCatchRange(wordRow, wcol, length - 1, ref start);
    }
    return res;
}

static void SetFontForMainCharracters(int wcol, Microsoft.Office.Interop.Word.Range r)
{
    if (wcol == 13)
    {
        r.Font.Size = 14;
        r.Font.Name = "SimSun";
    }
}