using System;
using System.Diagnostics;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace UserEditsFunctionality
{
    public partial class ThisAddIn
    {
        private Office.CommandBar CellContextMenu => Application.CommandBars["Cell"];
        private Office.CommandBar RowContextMenu => Application.CommandBars["Row"];

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ResetContextMenus();
            RegisterEvents();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ResetContextMenus();
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
            string sheetName = range.Worksheet.Name;
            if (!sheetName.EndsWith("_Model", StringComparison.CurrentCultureIgnoreCase))
            {
                ResetContextMenus();
                return;
            }

            CreateCustomContextMenu(CellContextMenu, range);
            CreateCustomContextMenu(RowContextMenu, range);
        }

        private void DeleteRow(Range range)
        {
            range.EntireRow.Delete();
        }

        private void ClassifyAsNewItem(Range range)
        {
            if (!CanClassifyAsNewItem(range))
                return;
            
            int rowCount = range.Rows.Count;

            //insert new empty rows
            Range baseRow = range.Worksheet.Rows[range.Row + rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                baseRow.Insert();
            }

            //copy values
            for (int i = 0; i < rowCount; i++)
            {
                Range newRow = ((Range)range.Worksheet.Rows[range.Row + rowCount + i]).EntireRow;
                Range oldRow = ((Range)range.Worksheet.Rows[range.Row + i]).EntireRow;

                newRow.Cell(1).Value = LabelAction.Addtition.GetDescription();
                newRow.Cell(2).Value = oldRow.Cell(2).Value;
                newRow.Cell(3).Value = oldRow.Cell(3).Value;
                newRow.Cell(4).Value = oldRow.Cell(4).Value;
                newRow.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(163, 255, 163));

                int lastNonEmptyColumn = oldRow.LastNonEmptyCell().Column;
                newRow.Cell(lastNonEmptyColumn).Value = oldRow.Cell(lastNonEmptyColumn).Value;

                oldRow.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(253, 157, 166));
                oldRow.Cell(1).Value = LabelAction.Discontinued.GetDescription();
                oldRow.Cell(lastNonEmptyColumn).Value = null;
            }
        }

        private void ClassifyAsLabelChange(Range range)
        {
            if (!CanClassifyAsLabelChange(range))
                return;

            foreach (Range row in range.Rows)
            {
                Range nonEmptyCells = row.EntireRow.Cells.SpecialCells(XlCellType.xlCellTypeConstants, 7);
                int aboveRowIndex = row.Row - range.Rows.Count;
                if(aboveRowIndex < 1)
                    continue;
                
                Range aboveRow = range.Worksheet.Rows[aboveRowIndex];

                foreach (Range cell in nonEmptyCells.Cells)
                {
                    aboveRow.Cell(cell.Column).Value = cell.Value;
                }

                aboveRow.Cell(1).Value = LabelAction.Change.GetDescription();
                aboveRow.Cell(3).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(163, 224, 255));
            }

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
            foreach (Range row in range.Rows)
            {
                string label = row.EntireRow.Cell(1).Value;
                if (string.IsNullOrEmpty(label) || !label.Equals(LabelAction.Addtition.GetDescription(),
                        StringComparison.CurrentCultureIgnoreCase))
                    return false;
            }
            
            return true;
        }

        private void ResetContextMenus()
        {
            CellContextMenu.Reset();
            RowContextMenu.Reset();
        }

        private void CreateCustomContextMenu(Office.CommandBar commandBar, Range range)
        {
            foreach (Office.CommandBarControl control in commandBar.Controls)
            {
                control.Delete(true);
            }

            var deleteRowMenuItem = commandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            deleteRowMenuItem.Caption = "Delete Row";
            ((CommandBarButton)deleteRowMenuItem).Click += (CommandBarButton deleteButton, ref bool cancelDelete) =>
            {
                DeleteRow(range);
            };

            var newItemMenuItem = commandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            newItemMenuItem.Caption = "Classify as New Item";
            newItemMenuItem.Enabled = CanClassifyAsNewItem(range);
            ((CommandBarButton)newItemMenuItem).Click += (CommandBarButton classifyNewButton, ref bool cancelNewItem) =>
            {
                ClassifyAsNewItem(range);
            };

            var labelChangeMenuItem = commandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            labelChangeMenuItem.Caption = "Classify as Label Change";
            labelChangeMenuItem.Enabled = CanClassifyAsLabelChange(range);
            ((CommandBarButton)labelChangeMenuItem).Click += (CommandBarButton classifyLabelChangeButton, ref bool cancelLabelChange) =>
            {
                ClassifyAsLabelChange(range);
            };
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
