Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Utils
Imports System.IO
Imports DevExpress.Office.Utils

Namespace Walkthrough_Creating_Table
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub richEditControl1_DocumentLoaded(ByVal sender As Object, ByVal e As EventArgs) Handles richEditControl1.DocumentLoaded
            CreateStyles()
        End Sub
        Private Sub btnCreateTable_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCreateTable.Click
            CreateTable()
            FillTable()
            ApplyHeadingStyle()
        End Sub

        Private Sub CreateTable()

            Dim doc As Document = richEditControl1.Document
            ' Clear out the document content
            doc.Delete(richEditControl1.Document.Range)
            ' Set up header information
            Dim pos As DocumentPosition = doc.Range.Start
            Dim rng As DocumentRange = doc.InsertSingleLineText(pos, "Directory Information from C:\")

            Dim cp_Header As CharacterProperties = doc.BeginUpdateCharacters(rng)
            cp_Header.FontName = "Verdana"
            cp_Header.FontSize = 16
            doc.EndUpdateCharacters(cp_Header)
            doc.Paragraphs.Insert(rng.End)
            doc.Paragraphs.Insert(rng.End)

            ' Add the table
            doc.Tables.Create(rng.End, 1, 3, AutoFitBehaviorType.AutoFitToWindow)
            ' Format the table
            Dim tbl As Table = doc.Tables(0)

            Try
                tbl.BeginUpdate()

                Dim cp_Tbl As CharacterProperties = doc.BeginUpdateCharacters(tbl.Range)
                cp_Tbl.FontSize = 8
                cp_Tbl.FontName = "Verdana"
                doc.EndUpdateCharacters(cp_Tbl)

                ' Insert header caption and format the columns
                doc.InsertSingleLineText(tbl(0, 0).Range.Start, "Name")
                doc.InsertSingleLineText(tbl(0, 1).Range.Start, "Size")
                Dim pp_HeadingSize As ParagraphProperties = doc.BeginUpdateParagraphs(tbl(0, 1).Range)
                pp_HeadingSize.Alignment = ParagraphAlignment.Right
                doc.EndUpdateParagraphs(pp_HeadingSize)

                doc.InsertSingleLineText(tbl(0, 2).Range.Start, "Modified")
                Dim pp_HeadingModified As ParagraphProperties = doc.BeginUpdateParagraphs(tbl(0, 2).Range)
                pp_HeadingModified.Alignment = ParagraphAlignment.Right
                doc.EndUpdateParagraphs(pp_HeadingModified)
                ' Apply a style to the table
                tbl.Style = doc.TableStyles("MyTableGridNumberEight")
                ' Specify right and left paddings equal to 0.08 inches for all cells in a table
                tbl.RightPadding = Units.InchesToDocumentsF(0.08F)
                tbl.LeftPadding = Units.InchesToDocumentsF(0.08F)
            Finally
                tbl.EndUpdate()
            End Try
        End Sub

        Private Sub CreateStyles()
            ' Define basic style
            Dim tStyleNormal As TableStyle = richEditControl1.Document.TableStyles.CreateNew()
            tStyleNormal.LineSpacingType = ParagraphLineSpacing.Single
            tStyleNormal.FontName = "Verdana"
            tStyleNormal.Alignment = ParagraphAlignment.Left
            tStyleNormal.Name = "MyTableGridNormal"
            richEditControl1.Document.TableStyles.Add(tStyleNormal)

            ' Define Grid Eight style
            Dim tStyleGrid8 As TableStyle = richEditControl1.Document.TableStyles.CreateNew()
            tStyleGrid8.Parent = tStyleNormal
            Dim borders As TableBorders = tStyleGrid8.TableBorders

            borders.Bottom.LineColor = Color.DarkBlue
            borders.Bottom.LineStyle = TableBorderLineStyle.Single
            borders.Bottom.LineThickness = 0.75F

            borders.Left.LineColor = Color.DarkBlue
            borders.Left.LineStyle = TableBorderLineStyle.Single
            borders.Left.LineThickness = 0.75F

            borders.Right.LineColor = Color.DarkBlue
            borders.Right.LineStyle = TableBorderLineStyle.Single
            borders.Right.LineThickness = 0.75F

            borders.Top.LineColor = Color.DarkBlue
            borders.Top.LineStyle = TableBorderLineStyle.Single
            borders.Top.LineThickness = 0.75F

            borders.InsideVerticalBorder.LineColor = Color.DarkBlue
            borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single
            borders.InsideVerticalBorder.LineThickness = 0.75F

            borders.InsideHorizontalBorder.LineColor = Color.DarkBlue
            borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single
            borders.InsideHorizontalBorder.LineThickness = 0.75F

            tStyleGrid8.CellBackgroundColor = Color.Transparent
            tStyleGrid8.Name = "MyTableGridNumberEight"
            richEditControl1.Document.TableStyles.Add(tStyleGrid8)

            ' Define Headings paragraph style
            Dim pStyleHeadings As ParagraphStyle = richEditControl1.Document.ParagraphStyles.CreateNew()
            pStyleHeadings.Bold = True
            pStyleHeadings.ForeColor = Color.White
            pStyleHeadings.Name = "My Headings Style"
            richEditControl1.Document.ParagraphStyles.Add(pStyleHeadings)
        End Sub

        Private Sub FillTable()
            ' Fill the table with data
            Dim doc As Document = richEditControl1.Document
            Dim tbl As Table = doc.Tables(0)
            Dim di As New DirectoryInfo("C:\")

            Try
                tbl.BeginUpdate()
                For Each fi As FileInfo In di.GetFiles()
                    Dim row As TableRow = tbl.Rows.Append()
                    Dim cell As TableCell = row.FirstCell
                    doc.InsertSingleLineText(cell.Range.Start, fi.Name)
                    doc.InsertSingleLineText(cell.Next.Range.Start, String.Format("{0:N0}", fi.Length))
                    doc.InsertSingleLineText(cell.Next.Next.Range.Start, String.Format("{0:g}", fi.LastWriteTime))
                Next fi
            Finally
                tbl.EndUpdate()
            End Try
        End Sub

        Private Sub ApplyHeadingStyle()
            Dim doc As Document = richEditControl1.Document
            Dim tbl As Table = doc.Tables(0)
            For Each cell As TableCell In tbl.Rows.First.Cells
                cell.BackgroundColor = Color.DarkBlue
            Next cell
            Dim pp_Headings As ParagraphProperties = doc.BeginUpdateParagraphs(tbl.Rows.First.Range)
            pp_Headings.Style = doc.ParagraphStyles("My Headings Style")
            doc.EndUpdateParagraphs(pp_Headings)
        End Sub
    End Class
End Namespace