using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
			
namespace WordFormsApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            Object oFalse = false;

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            oWord.ScreenUpdating = false;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

            ///////////////////////
            Word.Style style1 = oDoc.Styles.Add("MyHeading4");
            style1.set_BaseStyle("Heading 4"); //this is the ID of the style
            style1.Font.Size = 14;
            int num1 = style1.Creator;
            Word.Style style2 = oDoc.Styles.Add("MyHeading5");
            style2.set_BaseStyle("Heading 5"); //this is the ID of the style
            style2.Font.Bold = 1;
            int num2 = style1.Creator;
            /////////////////////////

            insertHtmlFile(ref oWord, ref oDoc, "c:/temp/word/merge.html", "Html Title", "Html Sub Title");

            insertFileLink(ref oWord, ref oDoc, "merge.pdf", "pdf title");

            /////////////////////////////////
            insertPageNumbers(ref oWord, ref oDoc, "Mission Name");

            ///////////////////////
            insertTableOfContents(ref oWord, ref oDoc);

           ///////////////////////
            saveAsPdf(ref oWord, ref oDoc, "c:/temp/word/test.pdf");

            oDoc.Close(ref oFalse, ref oMissing, ref oMissing);

            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

        }

        private void insertHtmlFile(ref Word._Application oWord, ref Word._Document oDoc, string fname, string title, string subTitle)
        {
            Object oTrue = true;
            Object oFalse = false;
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; // \endofdoc is a predefined bookmark
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            //Object styleHeading1 = "Heading 5";
            //Object styleHeading2 = "Heading 6";
            Object styleHeading1 = "MyHeading4";
            Object styleHeading2 = "MyHeading5";
         
            Word.Paragraph oHtmlTitle;
            Word.Paragraph oSectionTitle;
            oSectionTitle = oDoc.Content.Paragraphs.Add(ref oMissing);
            oSectionTitle.Range.Text = title;
            oSectionTitle.Range.Font.Bold = 1;
            oSectionTitle.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oSectionTitle.Range.set_Style(styleHeading1);
            oSectionTitle.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel4;
            oSectionTitle.Range.InsertParagraphAfter();
            
            oHtmlTitle = oDoc.Content.Paragraphs.Add(ref oMissing);
            oHtmlTitle.Range.Text = subTitle;
            oHtmlTitle.Range.Font.Bold = 1;
            oHtmlTitle.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oHtmlTitle.Range.set_Style(styleHeading2);
            oHtmlTitle.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel5;
            oHtmlTitle.Range.InsertParagraphAfter();

            Object oRngoBookMarkStart = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.Start;

            String oMergePath1 = "c:/temp/word/merge.html";
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertFile(oMergePath1, ref oMissing, ref oFalse, ref oFalse, ref oFalse);

            Object oRngoBookMarkEnd = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.Start;

            Word.Range rngBKMarkSelection = oDoc.Range(ref oRngoBookMarkStart, ref oRngoBookMarkEnd);
            //rngBKMarkSelection.set_Style(ref styleNormal);
            rngBKMarkSelection.Font.Shrink();
            rngBKMarkSelection.Font.Shrink();
            rngBKMarkSelection.Font.Shrink();
            rngBKMarkSelection.Font.Shrink();
            rngBKMarkSelection.Font.Shrink();

            //oWord.Selection.ClearFormatting();
        }

        private void insertFileLink(ref Word._Application oWord, ref Word._Document oDoc, string fname, string title)
        {
            object oEndOfDoc = "\\endofdoc"; // \endofdoc is a predefined bookmark
            object oMissing = System.Reflection.Missing.Value;
            Object oRange = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Object oAddress = fname;
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.Hyperlinks.Add(oRange, ref oAddress, ref oMissing, ref oMissing, title, ref oMissing);
            Object styleNormal = "Normal";
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.set_Style(styleNormal);
        }

        private void saveAsPdf(ref Word._Application oWord, ref Word._Document oDoc, string fname)
        {
            object oMissing = System.Reflection.Missing.Value;
            object outputFileName = fname;

            try
            {
                if (File.Exists(outputFileName.ToString()))
                {
                    File.Delete(outputFileName.ToString());
                }

                object fileFormat = Word.WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                oDoc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception e)
            {
            }
        }

        private void insertPageNumbers(ref Word._Application oWord, ref Word._Document oDoc, string name)
        {
            object oMissing = System.Reflection.Missing.Value;
            oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            //ENTERING A PARAGRAPH BREAK "ENTER"
            oWord.Selection.TypeParagraph();
            //INSERTING THE PAGE NUMBERS CENTRALLY ALIGNED IN THE PAGE FOOTER
            oWord.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oWord.ActiveWindow.Selection.Font.Name = "Arial";
            oWord.ActiveWindow.Selection.Font.Size = 8;
            oWord.ActiveWindow.Selection.TypeText(name);
            //INSERTING TAB CHARACTERS
            oWord.ActiveWindow.Selection.TypeText("\t");
            oWord.ActiveWindow.Selection.TypeText("\t");
            oWord.ActiveWindow.Selection.TypeText("Page ");
            Object CurrentPage = Word.WdFieldType.wdFieldPage;
            oWord.ActiveWindow.Selection.Fields.Add(oWord.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
            oWord.ActiveWindow.Selection.TypeText(" of ");
            Object TotalPages = Word.WdFieldType.wdFieldNumPages;
            oWord.ActiveWindow.Selection.Fields.Add(oWord.Selection.Range, ref TotalPages, ref oMissing, ref oMissing);
            //SETTING FOCUES BACK TO DOCUMENT
            oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        private void insertTableOfContents(ref Word._Application oWord, ref Word._Document oDoc)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; // \endofdoc is a predefined bookmark
            Object oTrue = true;
            Object oFalse = false;
            Object oUpperHeadingLevel = "4";
            Object oLowerHeadingLevel = "5";
            Object oTOCTableID = "TableOfContents";
            Word.Range rngTOC = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oDoc.TablesOfContents.Add(rngTOC, ref oTrue, ref oUpperHeadingLevel,
                               ref oLowerHeadingLevel, ref oMissing, ref oTOCTableID, ref oTrue,
                               ref oTrue, ref oMissing, ref oTrue, ref oTrue, ref oTrue);
        }

    }
}




/*
    ///////////////////////
    String oMergePath4 = "c:/temp/word/merge.png";
    oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InlineShapes.AddPicture(oMergePath4, ref oMissing, ref oMissing, ref oMissing);
    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
    wrdRng.InsertBreak(ref oPageBreak);
    ///////////////////////
    String oMergePath5 = "c:/temp/word/merge.docx";
    oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertFile(oMergePath5, ref oMissing, ref oFalse, ref oFalse, ref oFalse);
    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
    wrdRng.InsertBreak(ref oPageBreak);
*/
