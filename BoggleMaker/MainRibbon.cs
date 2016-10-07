using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BoggleMaker.Properties;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace BoggleMaker {
    public partial class MainRibbon {
        private static Random rand = new Random();
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e) {

        }

        private void InsertNew4_Click(object sender, RibbonControlEventArgs e) {
            Document document = Globals.ThisAddIn.Application.ActiveDocument;
            //typing header
            Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            currentSelection.TypeText("R&D IT's");
            try {
                var filePath = @"Resources\boggle_logo.png";
                currentSelection.InlineShapes.AddPicture(filePath);
            }
            catch (Exception ex) {
                currentSelection.TypeText(" Boggle");
            }
            //typing rules
            currentSelection.TypeText("\n\n1. Words must be 3+ letters\n\n"+
                                      "2. Words cannot be: names, hyphenated or abbreviated.\n\n" +
                                      "3. Write your words (linked or separated) in the table below:\n\n");
            //inserting table
            currentSelection = Globals.ThisAddIn.Application.Selection;
            document.Tables.Add(currentSelection.Range, 4, 4);
            Tables tables = document.Tables;
            Table boggleTable = tables[tables.Count];
            #region boggleTable styling
            boggleTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            boggleTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            boggleTable.Range.Font.Size = 36;
            boggleTable.Range.Font.Bold = 1;
            boggleTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            #endregion
            FillTableWithLetters(boggleTable);
            
        }

        private void FillTableWithLetters(Table boggleTable) {
            int maxCol = boggleTable.Columns.Count;
            int maxRow = boggleTable.Rows.Count;
            for (int col = 1; col <= maxCol; col++) {
                for (int row = 1; row <= maxRow; row++) {
                    Cell cell = boggleTable.Cell(row, col);
                    cell.Range.Text = GenRandomLetter();
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    if (row % 2 == 0 && col % 2 == 0)
                        cell.Range.Shading.BackgroundPatternColor = WdColor.wdColorGray20;
                    if (row % 2 == 1 && col % 2 == 1)
                        cell.Range.Shading.BackgroundPatternColor = WdColor.wdColorGray15;
                    cell.SetWidth(60, WdRulerStyle.wdAdjustNone);
                }
            }
        }

        private string GenRandomLetter() {
            string[] consonants = {
                "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N",
                "P", "Qu", "R", "S", "T", "V", "W", "X", "Y", "Z"
            };
            string[] vowels = {
                "A", "E", "I", "O", "U"
            };
            int chances = rand.Next(5);
            switch (chances) {
                case 1: case 2:
                    return vowels[rand.Next(5)];
                default:
                    return consonants[rand.Next(21)];
            }
        }
    }
}
