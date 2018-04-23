using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Utils;
using System.IO;

namespace Walkthrough_Creating_Table
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void richEditControl1_DocumentLoaded(object sender, EventArgs e)
        {
            CreateStyles();
        }
        private void btnCreateTable_Click(object sender, EventArgs e)
        {
            CreateTable();
            FillTable();
            ApplyHeadingStyle();
        }

        private void CreateTable() {

            Document doc = richEditControl1.Document;
            // Clear out the document content
            doc.Delete(richEditControl1.Document.Range);
            // Set up header information
            DocumentPosition pos = doc.Range.Start;
            DocumentRange rng = doc.InsertSingleLineText(pos, "Directory Information from C:\\");

            CharacterProperties cp_Header = doc.BeginUpdateCharacters(rng);
            cp_Header.FontName = "Verdana";
            cp_Header.FontSize = 16;
            doc.EndUpdateCharacters(cp_Header);
            doc.InsertParagraph(rng.End);
            doc.InsertParagraph(rng.End);
 
            // Add the table
            doc.Tables.Add(rng.End, 1, 3, AutoFitBehaviorType.AutoFitToWindow);
            // Format the table
            Table tbl = doc.Tables[0];

            try {
                tbl.BeginUpdate();

                CharacterProperties cp_Tbl = doc.BeginUpdateCharacters(tbl.Range);
                cp_Tbl.FontSize = 8;
                cp_Tbl.FontName = "Verdana";
                doc.EndUpdateCharacters(cp_Tbl);

                // Insert header caption and format the columns
                doc.InsertSingleLineText(tbl[0, 0].Range.Start, "Name");
                doc.InsertSingleLineText(tbl[0, 1].Range.Start, "Size");
                ParagraphProperties pp_HeadingSize = doc.BeginUpdateParagraphs(tbl[0, 1].Range);
                pp_HeadingSize.Alignment = ParagraphAlignment.Right;
                doc.EndUpdateParagraphs(pp_HeadingSize);

                doc.InsertSingleLineText(tbl[0, 2].Range.Start, "Modified");
                ParagraphProperties pp_HeadingModified = doc.BeginUpdateParagraphs(tbl[0, 2].Range);
                pp_HeadingModified.Alignment = ParagraphAlignment.Right;
                doc.EndUpdateParagraphs(pp_HeadingModified);
                // Apply a style to the table
                tbl.Style = doc.TableStyles["MyTableGridNumberEight"];
                // Specify right and left paddings equal to 0.08 inches for all cells in a table
                tbl.RightPadding = Units.InchesToDocumentsF(0.08f);
                tbl.LeftPadding = Units.InchesToDocumentsF(0.08f);
            }
            finally {
                tbl.EndUpdate();
            }
        }

        private void CreateStyles() {
            // Define basic style
            TableStyle tStyleNormal = richEditControl1.Document.TableStyles.CreateNew();
            tStyleNormal.LineSpacingType = ParagraphLineSpacing.Single;
            tStyleNormal.FontName = "Verdana";
            tStyleNormal.Alignment = ParagraphAlignment.Left;
            tStyleNormal.Name = "MyTableGridNormal";
            richEditControl1.Document.TableStyles.Add(tStyleNormal);

            // Define Grid Eight style
            TableStyle tStyleGrid8 = richEditControl1.Document.TableStyles.CreateNew();
            tStyleGrid8.Parent = tStyleNormal;
            TableBorders borders = tStyleGrid8.TableBorders;
            
            borders.Bottom.LineColor = Color.DarkBlue;
            borders.Bottom.LineStyle = TableBorderLineStyle.Single;
            borders.Bottom.LineThickness = 0.75f;
            
            borders.Left.LineColor = Color.DarkBlue;
            borders.Left.LineStyle = TableBorderLineStyle.Single;
            borders.Left.LineThickness = 0.75f;

            borders.Right.LineColor = Color.DarkBlue;
            borders.Right.LineStyle = TableBorderLineStyle.Single;
            borders.Right.LineThickness = 0.75f;

            borders.Top.LineColor = Color.DarkBlue;
            borders.Top.LineStyle = TableBorderLineStyle.Single;
            borders.Top.LineThickness = 0.75f;

            borders.InsideVerticalBorder.LineColor = Color.DarkBlue;
            borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single;
            borders.InsideVerticalBorder.LineThickness = 0.75f;

            borders.InsideHorizontalBorder.LineColor = Color.DarkBlue;
            borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single;
            borders.InsideHorizontalBorder.LineThickness = 0.75f;

            tStyleGrid8.CellBackgroundColor = Color.Transparent;
            tStyleGrid8.Name = "MyTableGridNumberEight";
            richEditControl1.Document.TableStyles.Add(tStyleGrid8);
        
            // Define Headings paragraph style
            ParagraphStyle pStyleHeadings = richEditControl1.Document.ParagraphStyles.CreateNew();
            pStyleHeadings.Bold = true;
            pStyleHeadings.ForeColor = Color.White;
            pStyleHeadings.Name = "My Headings Style";
            richEditControl1.Document.ParagraphStyles.Add(pStyleHeadings);
        }

        private void FillTable()
        {
            // Fill the table with data
            Document doc = richEditControl1.Document;
            Table tbl = doc.Tables[0];
            DirectoryInfo di = new DirectoryInfo("C:\\");

            try {
                tbl.BeginUpdate();
                foreach (FileInfo fi in di.GetFiles()) {
                    TableRow row = tbl.Rows.Append();
                    TableCell cell = row.FirstCell;
                    doc.InsertSingleLineText(cell.Range.Start, fi.Name);
                    doc.InsertSingleLineText(cell.Next.Range.Start,
                        String.Format("{0:N0}", fi.Length));
                    doc.InsertSingleLineText(cell.Next.Next.Range.Start,
                        String.Format("{0:g}", fi.LastWriteTime));
                }
            }
            finally {
                tbl.EndUpdate();
            }
        }

        private void ApplyHeadingStyle() {
            Document doc = richEditControl1.Document;
            Table tbl = doc.Tables[0];
            foreach (TableCell cell in tbl.Rows.First.Cells) {
                cell.BackgroundColor = Color.DarkBlue;
            }
            ParagraphProperties pp_Headings = doc.BeginUpdateParagraphs(tbl.Rows.First.Range);
            pp_Headings.Style = doc.ParagraphStyles["My Headings Style"];
            doc.EndUpdateParagraphs(pp_Headings);
        }
    }
}