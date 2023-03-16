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

using System;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Interaction;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Views;
using Point = System.Drawing.Point;
#if WINFORM
using System.Windows.Forms;
using RGFloat = System.Single;
using RGImage = System.Drawing.Image;
#else
using RGFloat = System.Double;
using RGImage = System.Windows.Media.ImageSource;
#endif // WINFORM

namespace unvell.ReoGrid.CellTypes;
#if WINFORM

/// <summary>
///     Represents an abstract base class for custom drop-down cell.
/// </summary>
public abstract class DropdownCell : CellBody
{
    /// <summary>
    ///     Get dropdown panel.
    /// </summary>
    public DropdownWindow DropdownPanel { get;  set; }

    /// <summary>
    ///     Determines whether or not to open the drop-down panel when user clicked inside cell.
    /// </summary>
    public virtual bool PullDownOnClick { get; set; } = true;

    private Size dropdownButtonSize = new(20, 20);

    /// <summary>
    ///     Get or set the drop-down button size.
    /// </summary>
    public virtual Size DropdownButtonSize
    {
        get => dropdownButtonSize;
        set => dropdownButtonSize = value;
    }

    private bool dropdownButtonAutoHeight = true;

    /// <summary>
    ///     Determines whether or not to adjust the height of drop-down button to fit entire cell.
    /// </summary>
    public virtual bool DropdownButtonAutoHeight
    {
        get => dropdownButtonAutoHeight;
        set
        {
            dropdownButtonAutoHeight = value;
            OnBoundsChanged();
        }
    }

    private Rectangle dropdownButtonRect = new(0, 0, 20, 20);

    /// <summary>
    ///     Get the drop-down button bounds.
    /// </summary>
    protected Rectangle DropdownButtonRect => dropdownButtonRect;

    /// <summary>
    ///     Get or set the control in drop-down panel.
    /// </summary>
    public virtual Control DropdownControl { get; set; }

    /// <summary>
    ///     Override method to handle the event when drop-down control lost focus.
    /// </summary>
    protected virtual void OnDropdownControlLostFocus()
    {
        PullUp();
    }

    public bool isDropdown;

    /// <summary>
    ///     Get or set whether the drop-down button is pressed. When this value is set to true, the drop-down panel will popped
    ///     up.
    /// </summary>
    public bool IsDropdown
    {
        get => isDropdown;
        set
        {
            if (isDropdown != value)
            {
                if (value)
                    PushDown();
                else
                    PullUp();
            }
        }
    }

    /// <summary>
    ///     Create custom drop-down cell instance.
    /// </summary>
    public DropdownCell()
    {
    }

    /// <summary>
    ///     Process boundary changes.
    /// </summary>
    public override void OnBoundsChanged()
    {
        dropdownButtonRect.Width = dropdownButtonSize.Width;

        if (dropdownButtonRect.Width > Bounds.Width)
            dropdownButtonRect.Width = Bounds.Width;
        else if (dropdownButtonRect.Width < 3) dropdownButtonRect.Width = 3;

        if (dropdownButtonAutoHeight)
            dropdownButtonRect.Height = Bounds.Height - 1;
        else
            dropdownButtonRect.Height = Math.Min(DropdownButtonSize.Height, Bounds.Height - 1);

        dropdownButtonRect.X = Bounds.Right - dropdownButtonRect.Width;

        var valign = ReoGridVerAlign.General;

        if (Cell != null && Cell.InnerStyle != null
                         && Cell.InnerStyle.HasStyle(PlainStyleFlag.VerticalAlign))
            valign = Cell.InnerStyle.VAlign;

        switch (valign)
        {
            case ReoGridVerAlign.Top:
                dropdownButtonRect.Y = 1;
                break;

            case ReoGridVerAlign.General:
            case ReoGridVerAlign.Bottom:
                dropdownButtonRect.Y = Bounds.Bottom - dropdownButtonRect.Height;
                break;

            case ReoGridVerAlign.Middle:
                dropdownButtonRect.Y = Bounds.Top + (Bounds.Height - dropdownButtonRect.Height) / 2 + 1;
                break;
        }
    }

    /// <summary>
    ///     Paint the dropdown button inside cell.
    /// </summary>
    /// <param name="dc">Platform no-associated drawing context instance.</param>
    public override void OnPaint(CellDrawingContext dc)
    {
        // call base to draw cell background and text
        base.OnPaint(dc);

        // draw button surface
        OnPaintDropdownButton(dc, dropdownButtonRect);
    }

    /// <summary>
    ///     Draw the drop-down button surface.
    /// </summary>
    /// <param name="dc">ReoGrid cross-platform drawing context.</param>
    /// <param name="buttonRect">Rectangle of drop-down button.</param>
    protected virtual void OnPaintDropdownButton(CellDrawingContext dc, Rectangle buttonRect)
    {
        if (Cell != null)
        {
            if (Cell.IsReadOnly)
                ControlPaint.DrawComboButton(dc.Graphics.PlatformGraphics, (System.Drawing.Rectangle)buttonRect,
                    ButtonState.Inactive);
            else
                ControlPaint.DrawComboButton(dc.Graphics.PlatformGraphics, (System.Drawing.Rectangle)buttonRect,
                    isDropdown ? ButtonState.Pushed : ButtonState.Normal);
        }
    }

    /// <summary>
    ///     Process when mouse button pressed inside cell.
    /// </summary>
    /// <param name="e"></param>
    /// <returns></returns>
    public override bool OnMouseDown(CellMouseEventArgs e)
    {
        if (PullDownOnClick || dropdownButtonRect.Contains(e.RelativePosition))
        {
            if (isDropdown)
                PullUp();
            else
                PushDown();

            return true;
        }

        return false;
    }

    /// <summary>
    ///     Handle event when mouse moving inside this cell body.
    /// </summary>
    /// <param name="e">Argument of mouse moving event.</param>
    /// <returns>True if event has been handled; Otherwise return false.</returns>
    public override bool OnMouseMove(CellMouseEventArgs e)
    {
        if (dropdownButtonRect.Contains(e.RelativePosition))
        {
            e.CursorStyle = CursorStyle.Hand;
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Handle event if cell has lost focus.
    /// </summary>
    public override void OnLostFocus()
    {
        PullUp();
    }

    /// <summary>
    ///     Event rasied when dropdown-panel is opened.
    /// </summary>
    public event EventHandler DropdownOpened;

    /// <summary>
    ///     Event raised when dropdown-panel is closed.
    /// </summary>
    public event EventHandler DropdownClosed;

    /// <summary>
    ///     Open dropdown panel when cell enter edit mode.
    /// </summary>
    /// <returns>True if edit operation is allowed; otherwise return false to abort edit.</returns>
    public override bool OnStartEdit()
    {
        PushDown();
        return false;
    }

    private Worksheet sheet;

    /// <summary>
    ///     Push down to open the dropdown panel.
    /// </summary>
    public virtual void PushDown()
    {
        if (Cell == null && Cell.Worksheet == null) return;

        if (Cell.IsReadOnly && DisableWhenCellReadonly) return;

        sheet = Cell == null ? null : Cell.Worksheet;
        
        if (sheet != null && DropdownControl != null
                          && CellsViewport.TryGetCellPositionToControl(sheet.ViewportController.FocusView,
                              Cell.InternalPos, out var p))
        {
            if (DropdownPanel == null)
            {
                DropdownPanel = new DropdownWindow(this);
                //dropdown.VisibleChanged += dropdown_VisibleChanged;

                //this.dropdownPanel.LostFocus -= DropdownControl_LostFocus;
                //this.dropdownPanel.OwnerItem = this.dropdownControl;
                DropdownPanel.VisibleChanged += DropdownPanel_VisibleChanged;
            }

            DropdownPanel.Width =
                Math.Max((int)Math.Round(Bounds.Width * sheet.renderScaleFactor), MinimumDropdownWidth);
            DropdownPanel.Height = DropdownPanelHeight;

            DropdownPanel.Show(sheet.workbook.ControlInstance,
                new Point((int)Math.Round(p.X), (int)Math.Round(p.Y + Bounds.Height * sheet.renderScaleFactor)));

            DropdownControl.Focus();

            isDropdown = true;
        }

        CallDropdownOpened();
    }

    public void CallDropdownOpened()
    {
        DropdownOpened?.Invoke(this, null);
    }

    public void DropdownPanel_VisibleChanged(object sender, EventArgs e)
    {
        OnDropdownControlLostFocus();
    }

    /// <summary>
    ///     Get or set height of dropdown-panel
    /// </summary>
    public virtual int DropdownPanelHeight { get; set; } = 200;

    /// <summary>
    ///     Minimum width of dropdown panel
    /// </summary>
    public virtual int MinimumDropdownWidth { get; set; } = 120;

    /// <summary>
    ///     Close condidate list
    /// </summary>
    public virtual void PullUp()
    {
        if (DropdownPanel != null)
        {
            DropdownPanel.Hide();

            isDropdown = false;

            if (sheet != null) sheet.RequestInvalidate();
        }

        if (DropdownClosed != null) DropdownClosed(this, null);
    }

    #region Dropdown Window

    /// <summary>
    ///     Prepresents dropdown window for dropdown cells.
    /// </summary>
#if WINFORM
    public class DropdownWindow : ToolStripDropDown
    {
        private readonly ToolStripControlHost controlHost;
        private readonly DropdownCell owner;

        /// <summary>
        ///     Create dropdown window instance.
        /// </summary>
        /// <param name="owner">The owner cell to this dropdown window.</param>
        public DropdownWindow(DropdownCell owner)
        {
            this.owner = owner;
            AutoSize = false;
            TabStop = true;

            Items.Add(controlHost = new ToolStripControlHost(this.owner.DropdownControl));

            controlHost.Margin = controlHost.Padding = new Padding(0);
            controlHost.AutoSize = false;
        }

        /// <summary>
        ///     Handle event when visible property changed.
        /// </summary>
        /// <param name="e">Arguments of visible changed event.</param>
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (!Visible)
            {
                owner.sheet.EndEdit(EndEditReason.Cancel);
            }
            else
            {
                if (owner.DropdownControl != null) BackColor = owner.DropdownControl.BackColor;
            }
        }

        /// <summary>
        ///     Handle event when size property changed.
        /// </summary>
        /// <param name="e">Arguments of size changed event.</param>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            if (controlHost != null)
                controlHost.Size = new System.Drawing.Size(ClientRectangle.Width - 2, ClientRectangle.Height - 2);
        }
    }
#elif WPF
		protected class DropdownWindow : System.Windows.Controls.Primitives.Popup
		{
			private DropdownCell owner;

			public DropdownWindow(DropdownCell owner)
			{
				this.owner = owner;
			}

			public void Hide()
			{
				this.IsOpen = false;
			}
		}
#endif // WPF

    #endregion // Dropdown Window
}

#endif // WINFORM