using Microsoft.Office.Interop;
using System.Reflection;

object missing = Missing.Value;
object endOfDoc = "\\endofdoc";

var applicationWord = new Microsoft.Office.Interop.Word.Application();
applicationWord.Visible = true;

var document = applicationWord.Documents.Add(ref missing, ref missing, ref missing, ref missing);

//Insert a paragraph at the end of the document.
object selectionRange = document.Bookmarks.get_Item(ref endOfDoc).Range;
var paragraph2 = document.Content.Paragraphs.Add(ref selectionRange);
paragraph2.Range.Text = "Heading 2";
paragraph2.Format.SpaceAfter = 6;
paragraph2.Range.InsertParagraphAfter();

//Insert another paragraph.
selectionRange = document.Bookmarks.get_Item(ref endOfDoc).Range;
var paragraph3 = document.Content.Paragraphs.Add(ref selectionRange);
paragraph3.Range.Text = "This is a sentence of normal text. Now here is a table:";
paragraph3.Range.Font.Bold = 0;
paragraph3.Format.SpaceAfter = 24;
paragraph3.Range.InsertParagraphAfter();

//Insert a 3 x 5 table, fill it with data, and make the first row bold and italic.
var wordRange = document.Bookmarks.get_Item(ref endOfDoc).Range;
var table = document.Tables.Add(wordRange, 3, 5, ref missing, ref missing);
table.Range.ParagraphFormat.SpaceAfter = 6;
int row, column;
string strText;
for (row = 1; row <= 3; row++)
    for (column = 1; column <= 5; column++)
    {
        strText = "r" + row + "c" + column;
        table.Cell(row, column).Range.Text = strText;
    }
table.Rows[1].Range.Font.Bold = 1;
table.Rows[1].Range.Font.Italic = 1;

Console.ReadLine();