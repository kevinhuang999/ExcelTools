namespace ExcelTools.Common;

using System.Drawing;
using System.Windows.Forms;

/// <summary>
/// Shared visual theme and tooltip helper.
///
/// Only colours, borders and button styling live here. Sizes and positions are
/// owned by <see cref="FormLayout"/> and are never modified by the theme: the
/// earlier version grew buttons in place, which pushed them past the right edge
/// and on top of their neighbours.
/// </summary>
internal static class UiTheme
{
    private const string FontFamily = "Microsoft YaHei UI";

    public static readonly Font HintFont = new(FontFamily, 8.25F, FontStyle.Regular, GraphicsUnit.Point);

    /// <summary>Excel-flavoured accent used for the primary action button.</summary>
    public static readonly Color Accent = Color.FromArgb(33, 115, 70);
    public static readonly Color Text = Color.FromArgb(45, 45, 45);
    public static readonly Color Muted = Color.FromArgb(122, 122, 122);
    public static readonly Color Surface = Color.FromArgb(250, 250, 250);
    public static readonly Color ButtonFace = Color.FromArgb(238, 238, 238);
    public static readonly Color Border = Color.FromArgb(206, 206, 206);

    private static ToolTip? _sharedTooltip;

    /// <summary>Recolour a form and give any plain controls a light touch-up.</summary>
    public static void Apply(Form form)
    {
        form.BackColor = Surface;
        form.ForeColor = Text;
        StyleChildren(form);
    }

    /// <summary>Attach a usage hint to a control using a shared ToolTip.</summary>
    public static void Tip(Control control, string text)
    {
        if (control == null || string.IsNullOrEmpty(text)) return;

        _sharedTooltip ??= new ToolTip
        {
            AutoPopDelay = 20000,
            InitialDelay = 350,
            ReshowDelay = 150,
            ShowAlways = true
        };
        _sharedTooltip.SetToolTip(control, text);
    }

    /// <summary>Accented action button. Size is owned by the caller.</summary>
    public static void StylePrimary(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
        button.BackColor = Accent;
        button.ForeColor = Color.White;
        button.FlatAppearance.BorderSize = 0;
        button.Tag = "styled";
    }

    /// <summary>Neutral button, used for Cancel and secondary actions.</summary>
    public static void StyleSecondary(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
        button.BackColor = ButtonFace;
        button.ForeColor = Text;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = Border;
        button.Tag = "styled";
    }

    private static void StyleChildren(Control parent)
    {
        foreach (Control control in parent.Controls)
        {
            switch (control)
            {
                case Button button when !"styled".Equals(button.Tag as string):
                    button.FlatStyle = FlatStyle.Flat;
                    button.Cursor = Cursors.Hand;
                    button.UseVisualStyleBackColor = false;
                    break;
                case TextBox textBox:
                    textBox.BackColor = Color.White;
                    break;
                case ProgressBar progress:
                    progress.Style = ProgressBarStyle.Continuous;
                    break;
            }

            if (control.HasChildren)
                StyleChildren(control);
        }
    }
}
