namespace ExcelTools.Common;

using System.Drawing;
using System.Windows.Forms;

/// <summary>
/// Declarative layout builder for the tool dialogs.
///
/// The earlier hand-positioned layouts were cramped and overlapped: buttons
/// were grown in place and pushed past the right edge, while the usage line
/// landed on top of the button row. This builder derives every position from
/// one set of metrics and sizes the dialog to the widest row, so margins,
/// spacing, the hint block and the button row can never collide.
///
/// Only presentation is affected - control order and the click flow are the
/// same as before.
/// </summary>
internal sealed class FormLayout
{
    public const int Margin = 20;
    public const int LabelWidth = 118;
    public const int Gap = 10;
    public const int DefaultBoxWidth = 150;
    public const int RowHeight = 34;
    public const int ButtonHeight = 32;
    public const int ButtonWidth = 100;
    public const int ButtonGap = 12;
    public const int MinWidth = 380;

    private readonly Form _form;
    private readonly string _usage;
    private readonly List<Control> _rowControls = new();
    private int _y;
    private int _width;

    public FormLayout(Form form, string usage)
    {
        _form = form;
        _usage = usage;
        _y = Margin + 4;
        _width = MinWidth;
    }

    private void Require(int rightEdge)
        => _width = Math.Max(_width, rightEdge + Margin);

    private Label NewCaption(string text, int top, int width)
    {
        var caption = new Label
        {
            Text = text,
            Location = new Point(Margin, top),
            Size = new Size(width, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = UiTheme.Text,
            AutoSize = false
        };
        _form.Controls.Add(caption);
        return caption;
    }

    /// <summary>Label on the left, single-line textbox on the right.</summary>
    public TextBox AddField(string label, string initialValue, int boxWidth = DefaultBoxWidth)
    {
        var caption = NewCaption(label, _y + 4, LabelWidth);

        var box = new TextBox
        {
            Location = new Point(Margin + LabelWidth + Gap, _y),
            Width = boxWidth,
            Text = initialValue,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _form.Controls.Add(box);

        Require(Margin + LabelWidth + Gap + boxWidth);
        _y += RowHeight;

        UiTheme.Tip(caption, label);
        UiTheme.Tip(box, label);
        return box;
    }

    /// <summary>Field row followed by a small button on the same line.</summary>
    public (TextBox Box, Button Button) AddFieldWithButton(
        string label, string initialValue, string buttonText, int boxWidth = 200)
    {
        var box = AddField(label, initialValue, boxWidth);

        var button = new Button
        {
            Text = buttonText,
            Location = new Point(Margin + LabelWidth + Gap + boxWidth + Gap, box.Top),
            Size = new Size(ButtonWidth, 26)
        };
        UiTheme.StyleSecondary(button);
        _form.Controls.Add(button);

        Require(Margin + LabelWidth + Gap + boxWidth + Gap + ButtonWidth);
        _rowControls.Add(button);
        return (box, button);
    }

    /// <summary>Block label with a multi-line text area underneath.</summary>
    public RichTextBox AddMultilineField(string label, int height, int width = 360)
    {
        NewCaption(label, _y, width);
        _y += 24;

        var box = new RichTextBox
        {
            Location = new Point(Margin, _y),
            Size = new Size(width, height),
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _form.Controls.Add(box);

        Require(Margin + width);
        _y += height + 10;
        return box;
    }

    /// <summary>Full-width progress bar.</summary>
    public ProgressBar AddProgressBar()
    {
        int width = 320;
        var bar = new ProgressBar
        {
            Location = new Point(Margin, _y + 4),
            Size = new Size(width, 18),
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous
        };
        _form.Controls.Add(bar);

        Require(Margin + width);
        _y += 28;
        return bar;
    }

    /// <summary>Small muted status line, e.g. "12/480".</summary>
    public Label AddStatusLabel()
    {
        var label = new Label
        {
            Text = string.Empty,
            Location = new Point(Margin, _y),
            Size = new Size(200, 18),
            ForeColor = UiTheme.Muted,
            Font = UiTheme.HintFont,
            AutoSize = false
        };
        _form.Controls.Add(label);
        _y += 22;
        return label;
    }

    /// <summary>
    /// Emit the usage line and the button row, size the dialog, and return the
    /// created buttons in declaration order.
    /// </summary>
    public Button[] Finish(params (string Text, bool Primary)[] buttons)
    {
        // Usage block, always below the last input row.
        int usageTop = _y + 6;
        int lines = Math.Max(1, _usage.Split('\n').Length);
        int usageHeight = 16 * lines + 6;

        var hint = new Label
        {
            Text = _usage,
            Location = new Point(Margin, usageTop),
            Size = new Size(_width - Margin * 2, usageHeight),
            ForeColor = UiTheme.Muted,
            Font = UiTheme.HintFont,
            AutoSize = false
        };
        _form.Controls.Add(hint);

        // Button row, right aligned, always clear of the usage block.
        int buttonsTop = usageTop + usageHeight + 14;
        var created = new Button[buttons.Length];

        int totalWidth = 0;
        for (int i = 0; i < buttons.Length; i++)
            totalWidth += (buttons[i].Primary ? 110 : 92) + (i > 0 ? ButtonGap : 0);

        int x = _width - Margin - totalWidth;
        for (int i = 0; i < buttons.Length; i++)
        {
            int width = buttons[i].Primary ? 110 : 92;
            var button = new Button
            {
                Text = buttons[i].Text,
                Location = new Point(x, buttonsTop),
                Size = new Size(width, ButtonHeight)
            };

            if (buttons[i].Primary) UiTheme.StylePrimary(button);
            else UiTheme.StyleSecondary(button);

            _form.Controls.Add(button);
            created[i] = button;
            x += width + ButtonGap;
        }

        _form.ClientSize = new Size(_width, buttonsTop + ButtonHeight + Margin);
        return created;
    }
}
