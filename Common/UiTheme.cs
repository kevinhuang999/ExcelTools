namespace ExcelTools.Common;

using System.Drawing;
using System.Windows.Forms;

/// <summary>
/// Shared visual theme for every dialog.
///
/// The original add-in used stock WinForms styling: default grey background,
/// plain raised buttons, no hints. Only the look is modernised here.
///
/// IMPORTANT: this deliberately does NOT assign Form.Font. Changing the form
/// font under AutoScaleMode.Font makes WinForms rescale every child control,
/// which shifts and overlaps the existing layout. Control order, positions,
/// labels and the click flow must stay exactly as before, so only colours,
/// button styling, borders and an appended usage line are touched.
/// </summary>
internal static class UiTheme
{
    private const string FontFamily = "Microsoft YaHei UI";

    public static readonly Font HintFont = new(FontFamily, 8.25F, FontStyle.Regular, GraphicsUnit.Point);

    /// <summary>Excel-flavoured accent used for the primary action button.</summary>
    public static readonly Color Accent = Color.FromArgb(33, 115, 70);
    public static readonly Color AccentHover = Color.FromArgb(43, 145, 90);
    public static readonly Color Text = Color.FromArgb(45, 45, 45);
    public static readonly Color Muted = Color.FromArgb(122, 122, 122);
    public static readonly Color Surface = Color.FromArgb(250, 250, 250);
    public static readonly Color ButtonFace = Color.FromArgb(238, 238, 238);
    public static readonly Color Border = Color.FromArgb(206, 206, 206);

    /// <summary>Accented action button (already themed; Apply keeps it as is).</summary>
    public static Button PrimaryButton(string text)
    {
        var button = new Button
        {
            Text = text,
            Size = new Size(110, 30),
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Accent,
            ForeColor = Color.White
        };
        button.FlatAppearance.BorderSize = 0;
        return button;
    }

    /// <summary>Neutral button, used for Cancel.</summary>
    public static Button SecondaryButton(string text)
    {
        var button = new Button
        {
            Text = text,
            Size = new Size(95, 30),
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = ButtonFace,
            ForeColor = Text
        };
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = Border;
        return button;
    }

    public static Label FieldLabel(string text)
        => new() { Text = text, ForeColor = Text, AutoSize = true };

    /// <summary>Muted explanation label.</summary>
    public static Label HintLabel(string text)
        => new() { Text = text, Font = HintFont, ForeColor = Muted, AutoSize = false };

    /// <summary>
    /// Apply the theme to a form, append an optional usage line and attach
    /// tooltips. Call as the last statement of the form constructor.
    /// </summary>
    public static void Apply(Form form, string? usage = null, params (Control Control, string Tip)[] tips)
    {
        form.BackColor = Surface;
        form.ForeColor = Text;

        StyleChildren(form);
        AppendUsageLine(form, usage);

        if (tips.Length > 0)
        {
            var tooltip = new ToolTip
            {
                AutoPopDelay = 20000,
                InitialDelay = 350,
                ReshowDelay = 150,
                ShowAlways = true
            };
            foreach (var (control, tip) in tips)
            {
                if (control != null && !string.IsNullOrEmpty(tip))
                    tooltip.SetToolTip(control, tip);
            }
        }
    }

    /// <summary>Add a muted explanation line below the existing content.</summary>
    private static void AppendUsageLine(Form form, string? usage)
    {
        if (string.IsNullOrEmpty(usage)) return;

        int top = form.ClientSize.Height + 6;
        form.ClientSize = new Size(form.ClientSize.Width, form.ClientSize.Height + 46);

        var hint = new Label
        {
            Text = usage,
            Font = HintFont,
            ForeColor = Muted,
            Location = new Point(18, top),
            Size = new Size(form.ClientSize.Width - 36, 38),
            AutoSize = false
        };
        form.Controls.Add(hint);
    }

    private static void StyleChildren(Control parent)
    {
        foreach (Control control in parent.Controls)
        {
            switch (control)
            {
                case Button button:
                    StyleButton(button);
                    break;
                case Label label:
                    // Only recolour existing labels; never re-metric them.
                    if (label.ForeColor == SystemColors.ControlText)
                        label.ForeColor = Text;
                    break;
                case TextBox textBox:
                    textBox.BackColor = Color.White;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                    break;
                case RichTextBox richText:
                    richText.BackColor = Color.White;
                    richText.BorderStyle = BorderStyle.FixedSingle;
                    break;
                case ProgressBar progress:
                    progress.Style = ProgressBarStyle.Continuous;
                    break;
            }

            if (control.HasChildren)
                StyleChildren(control);
        }
    }

    /// <summary>
    /// Cancel-style buttons stay neutral; every other button becomes the
    /// accented primary action. This mirrors the intent of the original dialogs
    /// without changing which button does what.
    /// </summary>
    private static void StyleButton(Button button)
    {
        string caption = button.Text.Replace("&", string.Empty).Trim();
        bool isCancel = caption.Equals("Cancel", StringComparison.OrdinalIgnoreCase);

        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;

        if (isCancel)
        {
            button.BackColor = ButtonFace;
            button.ForeColor = Text;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = Border;
        }
        else
        {
            button.BackColor = Accent;
            button.ForeColor = Color.White;
            button.FlatAppearance.BorderSize = 0;
        }

        // Give every action button a consistent, comfortable size without
        // moving it: only grow if it was smaller than the target.
        if (button.Width < 95) button.Width = 95;
        if (button.Height < 30) button.Height = 30;
    }
}
