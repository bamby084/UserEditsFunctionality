using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace UserEditsFunctionality
{
    public partial class ThisAddIn
    {
        private Office.CommandBar ContextMenu => Application.CommandBars["Cell"];

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ContextMenu.Reset();
            RegisterEvents();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ContextMenu.Reset();
            UnRegisterEvents();
        }

        private void RegisterEvents()
        {
            this.Application.SheetBeforeRightClick += OnDisplayContextMenu;
        }

        private void UnRegisterEvents()
        {
            this.Application.SheetBeforeRightClick -= OnDisplayContextMenu;
        }

        private void OnDisplayContextMenu(object sheet, Range range, ref bool cancel)
        {
            foreach (Office.CommandBarControl control in ContextMenu.Controls)
            {
                control.Delete(true);
            }

            var deleteRowMenuItem = ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            deleteRowMenuItem.Caption = "Delete Row";
            ((CommandBarButton)deleteRowMenuItem).Click += (CommandBarButton deleteButton, ref bool cancelDelete) =>
            {
                DeleteRow(range);
            };

            var newItemMenuItem = ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            newItemMenuItem.Caption = "Classify as New Item";
            newItemMenuItem.Enabled = CanClassifyAsNewItem(range);
            ((CommandBarButton)newItemMenuItem).Click += (CommandBarButton classifyNewButton, ref bool cancelNewItem) =>
            {
                ClassifyAsNewItem(range);
            };

            var labelChangeMenuItem = ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            labelChangeMenuItem.Caption = "Classify as Label Change";
            labelChangeMenuItem.Enabled = CanClassifyAsLabelChange(range);
            ((CommandBarButton)labelChangeMenuItem).Click += (CommandBarButton classifyLabelChangeButton, ref bool cancelLabelChange) =>
            {
                ClassifyAsLabelChange(range);
            };
        }

        private void DeleteRow(Range range)
        {
            range.EntireRow.Delete();
        }

        private void ClassifyAsNewItem(Range range)
        {
            if (!CanClassifyAsNewItem(range))
                return;

            Range row = range.Worksheet.Rows[range.Row + 1];
            row.Insert();

            Range newRow = range.Worksheet.Rows[range.Row + 1];
            range.EntireRow.Copy(newRow);

            newRow.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(163, 255, 163));
            newRow.Cell(1).Value = LabelAction.Addtition.GetDescription();
            int lastCellColumn = newRow.LastNonEmptyCell().Column;

            for (int i = 5; i < lastCellColumn; i++)
            {
                newRow.Cell(i).Value = null;
            }

            range.EntireRow.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(253, 157, 166));
            range.EntireRow.Cell(1).Value = LabelAction.Discontinued.GetDescription();
            range.EntireRow.Cell(lastCellColumn).Value = null;
        }

        private void ClassifyAsLabelChange(Range range)
        {
            if (!CanClassifyAsLabelChange(range))
                return;

            Range nonEmptyCells = range.EntireRow.Cells.SpecialCells(XlCellType.xlCellTypeConstants, 7);
            Range rowAbove = range.Worksheet.Rows[range.Row - 1];

            foreach (Range cell in nonEmptyCells.Cells)
            {
                rowAbove.Cell(cell.Column).Value = cell.Value;
            }

            rowAbove.Cell(1).Value = LabelAction.Change.GetDescription();
            rowAbove.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(163, 224, 255));
            range.EntireRow.Delete();
        }

        private bool CanClassifyAsNewItem(Range range)
        {
            foreach (Range row in range.Rows)
            {
                string label = row.EntireRow.Cell(1).Value;
                if (label == null || !label.Equals(LabelAction.Change.GetDescription(),
                        StringComparison.CurrentCultureIgnoreCase))

                    return false;
            }

            return true;
        }

        private bool CanClassifyAsLabelChange(Range range)
        {
            string label = range.EntireRow.Cell(1).Value;
            if (string.IsNullOrEmpty(label) || !label.Equals(LabelAction.Addtition.GetDescription(),
                    StringComparison.CurrentCultureIgnoreCase) || range.Row == 1)
                return false;

            return true;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
