/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2023 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2023 unvell inc. All rights reserved.
 * 
 ****************************************************************************/

#if WINFORM

using System;

using System.Drawing;
using System.Windows.Forms;

using RGRectF = System.Drawing.RectangleF;
using unvell.ReoGrid.Interaction;

namespace unvell.ReoGrid.WinForm
{
    internal class ColumnFilterContextMenu : ContextMenuStrip
    {
        internal ToolStripItem SortAZItem { get; set; }
        internal ToolStripItem SortZAItem { get; set; }

        internal CheckedListBox CheckedListBox { get; set; }

        internal Button OkButton { get; set; }
        internal Button CancelButton { get; set; }

        public ColumnFilterContextMenu()
        {
            AutoSize = false;
            Width = 240;
            Height = 340;

            Items.Add(SortAZItem = new ToolStripMenuItem(LanguageResource.Filter_SortAtoZ));
            Items.Add(SortZAItem = new ToolStripMenuItem(LanguageResource.Filter_SortZtoA));
            Items.Add(new ToolStripSeparator());

            Items.Add(new ToolStripControlHost(
                CheckedListBox = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    TabStop = false,
                    CheckOnClick = true,
                })
            {
                AutoSize = false,
                Width = 200,
                Height = 240,
            });

            CheckedListBox.ItemCheck += checkedListBox_ItemCheck;

            var hardFilterPanel = new Panel
            {
                Padding = new Padding(0, 4, 0, 4),
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
            };

            hardFilterPanel.Controls.Add(OkButton = new Button
            {
                Text = LanguageResource.Button_OK,
                Dock = DockStyle.Right,
            });
            hardFilterPanel.Controls.Add(new Splitter
            {
                Enabled = false,
                Width = 4,
                Dock = DockStyle.Right,
            });
            hardFilterPanel.Controls.Add(CancelButton = new Button
            {
                Text = LanguageResource.Button_Cancel,
                Dock = DockStyle.Right,
            });

            Items.Add(new ToolStripControlHost(hardFilterPanel)
            {
                AutoSize = false,
                Width = 200,
                Height = 30,
            });
        }

        internal int SelectedCount { get; set; }

        private bool inEventProcess = false;

        void checkedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (inEventProcess) return;

            inEventProcess = true;

            // 'Select All'
            if (e.Index == 0)
            {
                if (CheckedListBox.GetItemCheckState(0) != e.NewValue)
                {
                    for (int i = 1; i < CheckedListBox.Items.Count; i++)
                    {
                        CheckedListBox.SetItemChecked(i, e.NewValue == CheckState.Checked);
                        SelectedCount = CheckedListBox.Items.Count - 1;
                    }

                    if (e.NewValue == CheckState.Checked)
                    {
                        SelectedCount = CheckedListBox.Items.Count - 1;
                    }
                    else
                    {
                        SelectedCount = 0;
                    }
                }
            }
            // others else 'Select All'
            else if (e.NewValue != CheckedListBox.GetItemCheckState(0))
            {
                CheckedListBox.SetItemCheckState(0, CheckState.Indeterminate);

                if (e.NewValue != CheckState.Checked)
                {
                    SelectedCount--;

                    if (SelectedCount <= 0)
                    {
                        CheckedListBox.SetItemChecked(0, false);
                        SelectedCount = 0;
                    }
                }
                else
                {
                    SelectedCount++;

                    if (SelectedCount >= CheckedListBox.Items.Count - 1)
                    {
                        CheckedListBox.SetItemChecked(0, true);
                    }
                }
            }

            inEventProcess = false;
        }

        internal static void ShowFilterPanel(Data.AutoColumnFilter.AutoColumnFilterBody headerBody, Point point)
        {
            if (headerBody.ColumnHeader == null || headerBody.ColumnHeader.Worksheet == null) return;

            var worksheet = headerBody.ColumnHeader.Worksheet;
            if (worksheet == null) return;

            RGRectF headerRect = Views.ColumnHeaderView.GetColHeaderBounds(worksheet, headerBody.ColumnHeader.Index, point);
            if (headerRect.Width == 0 || headerRect.Height == 0) return;

            RGRectF buttonRect = headerBody.GetColumnFilterButtonRect(headerRect.Size);

            if (headerBody.ContextMenuStrip == null)
            {
                var filterPanel = new ColumnFilterContextMenu();

                filterPanel.SortAZItem.Click += (s, e) =>
                {
                    try
                    {
                        worksheet.SortColumn(headerBody.ColumnHeader.Index, 
                            new RangePosition(0, headerBody.autoFilter.ApplyRange.Col, worksheet.MaxContentRow + 1, headerBody.autoFilter.ApplyRange.Cols),
                            SortOrder.Ascending);
                    }
                    catch (Exception ex)
                    {
                        worksheet.NotifyExceptionHappen(ex);
                    }
                };

                filterPanel.SortZAItem.Click += (s, e) =>
                {
                    try
                    {
                        worksheet.SortColumn(headerBody.ColumnHeader.Index, 
                            new RangePosition(0, headerBody.autoFilter.ApplyRange.Col, worksheet.MaxContentRow + 1, headerBody.autoFilter.ApplyRange.Cols),
                            SortOrder.Descending);
                    }
                    catch (Exception ex)
                    {
                        worksheet.NotifyExceptionHappen(ex);
                    }
                };

                filterPanel.OkButton.Click += (s, e) =>
                {
                    if (filterPanel.CheckedListBox.GetItemCheckState(0) == CheckState.Checked)
                    {
                        headerBody.IsSelectAll = true;
                    }
                    else
                    {
                        headerBody.IsSelectAll = false;
                        headerBody.selectedTextItems.Clear();

                        for (int i = 1; i < filterPanel.CheckedListBox.Items.Count; i++)
                        {
                            if (filterPanel.CheckedListBox.GetItemChecked(i))
                            {
                                headerBody.selectedTextItems.Add(Convert.ToString(filterPanel.CheckedListBox.Items[i]));
                            }
                        }
                    }

                    headerBody.autoFilter.Apply();
                    filterPanel.Hide();
                };

                filterPanel.CancelButton.Click += (s, e) => filterPanel.Hide();

                headerBody.ContextMenuStrip = filterPanel;
            }

            if (headerBody.ContextMenuStrip != null)
            {
                if (headerBody.ContextMenuStrip is ColumnFilterContextMenu filterPanel)
                {
                    if (headerBody.DataDirty)
                    {
                        // todo: keep select status for every items before clear
                        filterPanel.CheckedListBox.Items.Clear();

                        filterPanel.CheckedListBox.Items.Add(LanguageResource.Filter_SelectAll);
                        filterPanel.CheckedListBox.SetItemChecked(0, true);

                        try
                        {
                            headerBody.ColumnHeader.Worksheet.ControlAdapter.ChangeCursor(CursorStyle.Busy);

                            var items = headerBody.GetDistinctItems();
                            foreach (string item in items)
                            {
                                filterPanel.CheckedListBox.Items.Add(item);

                                if (headerBody.IsSelectAll)
                                {
                                    filterPanel.CheckedListBox.SetItemChecked(filterPanel.CheckedListBox.Items.Count - 1, true);
                                }
                                else
                                {
                                    filterPanel.CheckedListBox.SetItemChecked(filterPanel.CheckedListBox.Items.Count - 1,
                                        headerBody.selectedTextItems.Contains(item));
                                }
                            }
                        }
                        finally
                        {
                            headerBody.ColumnHeader.Worksheet.ControlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
                        }

                        filterPanel.SelectedCount = filterPanel.CheckedListBox.Items.Count - 1;

                        headerBody.DataDirty = false;

                        headerBody.IsSelectAll = true;
                    }
                }

                var pp = new Graphics.Point(headerRect.Right - 240, buttonRect.Bottom + 1);

                pp = worksheet.ControlAdapter.PointToScreen(pp);

                headerBody.ContextMenuStrip.Show(Point.Round(pp));
            }
        }
    }
}

#endif // WINFORM
