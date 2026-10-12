namespace ExcelTools.Forms;

using ExcelTools.Common;

/// <summary>
/// Progress window. TopMost, with a settable progress value and message.
/// </summary>
public class ProgressForm : Form
{
    private readonly ProgressBar _progressBar;
    private readonly Label _lblMessage;

    public ProgressForm()
    {
        this.Text = "Working";
        this.Size = new Size(380, 110);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ShowInTaskbar = false;

        _progressBar = new ProgressBar
        {
            Location = new Point(15, 25),
            Size = new Size(340, 24),
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous
        };

        _lblMessage = new Label
        {
            Location = new Point(15, 60),
            Size = new Size(340, 25),
            Text = "Working, please wait...",
            TextAlign = ContentAlignment.MiddleLeft
        };

        this.Controls.Add(_progressBar);
        this.Controls.Add(_lblMessage);

        UiTheme.Apply(this);
    }

    public void SetProgress(int percent)
    {
        try
        {
            if (this.InvokeRequired)
                this.Invoke(() => _progressBar.Value = Math.Clamp(percent, 0, 100));
            else
                _progressBar.Value = Math.Clamp(percent, 0, 100);
        }
        catch { }
    }

    public void SetMessage(string msg)
    {
        try
        {
            if (this.InvokeRequired)
                this.Invoke(() => _lblMessage.Text = msg);
            else
                _lblMessage.Text = msg;
        }
        catch { }
    }
}