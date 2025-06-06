using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

static class P
{
    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);

    [DllImport("user32.dll")]
    internal static extern bool DestroyIcon(IntPtr handle);

    private static NotifyIcon icon;
    private static int currentWeek;
    private static Icon weekIcon;

    [STAThread]
    static void Main()
    {
        Application.SetCompatibleTextRenderingDefault(false);
        Application.EnableVisualStyles();

        SetProcessDpiAwarenessContext(new IntPtr(-4));

        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer() { Interval = 10000, Enabled = true };
        timer.Tick += delegate { UpdateIcon(); };
        currentWeek = GetWeek();
        weekIcon = GetIcon(currentWeek);
        icon = new NotifyIcon()
        {
            Icon = weekIcon,
            Visible = true,
            ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("About", delegate { MessageBox.Show(Application.ProductName); }),
                new MenuItem("Open web site...", delegate { Process.Start("https://voltura.github.io/WeekNumberLite2Plus"); }),
                new MenuItem("Save Icon...", delegate {
                    using (FileStream fs = new FileStream(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WeekIcon.ico"), FileMode.Create, FileAccess.Write))
                    {
                        weekIcon.Save(fs);
                    }
                    MessageBox.Show("Icon saved to desktop.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }),
                new MenuItem("Start with Windows", delegate(object sender, EventArgs e) {
                        StartWithWindows = !((MenuItem)sender).Checked;
                        ((MenuItem)sender).Checked = StartWithWindows;
                    }) { Checked = StartWithWindows },
                new MenuItem("Exit", delegate { Application.Exit(); })
            })
        };

        UpdateIcon();
        Application.Run(new ApplicationContext());
    }

    private static int GetWeek()
    {
        DateTime now = DateTime.Now;
        int week = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

        return (week == 53 && !(Pp(now.Year) == 4 || Pp(now.Year - 1) == 3)) ? 1 : week;
    }

    private static int Pp(int year)
    {
        return (int)(year + Math.Floor(year / 4.0) - Math.Floor(year / 100.0) + Math.Floor(year / 400.0)) % 7;
    }

    private static void UpdateIcon()
    {
        try
        {
            int week = GetWeek();
            icon.Text = (Thread.CurrentThread.CurrentUICulture.LCID == 1053 ? "Vecka " : "Week ") + week + "\r\n" + DateTime.Now.ToString("yyyy-MM-dd");

            if (week != currentWeek)
            {
                Icon prevIcon = icon.Icon;

                weekIcon = GetIcon(week);
                icon.Icon = weekIcon;

                if (prevIcon != null)
                {
                    DestroyIcon(prevIcon.Handle);
                    prevIcon.Dispose();
                }

                currentWeek = week;
            }
        }
        catch { }
    }

    private static readonly int[] sizes = { 16, 32, 48, 64, 128, 256, 512 };

    internal static Icon GetIcon(int week)
    {
        using (MemoryStream iconStream = new MemoryStream())
        using (BinaryWriter writer = new BinaryWriter(iconStream))
        {
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write((ushort)sizes.Length);

            int imageOffset = 6 + (16 * sizes.Length);
            MemoryStream[] imageStreams = new MemoryStream[sizes.Length];

            for (int i = 0, size; i < sizes.Length; i++)
            {
                size = sizes[i];
                imageStreams[i] = new MemoryStream();

                using (Bitmap bmp = new Bitmap(size, size))
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                    DrawBackground(g, size);
                    DrawWeekText(g, size, week);

                    bmp.Save(imageStreams[i], ImageFormat.Png);
                }

                writer.Write((byte)(size >= 256 ? 0 : size));
                writer.Write((byte)(size >= 256 ? 0 : size));
                writer.Write((byte)0);
                writer.Write((byte)0);
                writer.Write((ushort)1);
                writer.Write((ushort)32);
                writer.Write((int)imageStreams[i].Length);
                writer.Write((uint)imageOffset);

                imageOffset += (int)imageStreams[i].Length;
            }

            foreach (var stream in imageStreams)
            {
                writer.Write(stream.ToArray());
                stream.Dispose();
            }

            Array.Clear(imageStreams, 0, imageStreams.Length);

            writer.Flush();
            iconStream.Position = 0;

            return new Icon(iconStream);
        }
    }

    private static void DrawBackground(Graphics g, int size)
    {
        float inset = size * 0.03125f;

        using (SolidBrush bg = new SolidBrush(Color.Black))
        using (SolidBrush fg = new SolidBrush(Color.LightGray))
        using (Pen border = new Pen(Color.LightGray, inset * 2))
        {
            g.FillRectangle(bg, inset, inset, size - inset, size - inset);
            g.DrawRectangle(border, inset, inset, size - (inset * 2), size - (inset * 2));
            g.FillRectangle(fg, size * 0.15625f, inset / 2, inset * 3, inset * 5);
            g.FillRectangle(fg, size * 0.75f, inset / 2, inset * 3, inset * 5);
        }
    }

    private static void DrawWeekText(Graphics g, int size, int week)
    {
        float fontSize = size * 0.78125f;

        using (Font font = new Font(FontFamily.GenericMonospace, fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
        using (Brush brush = new SolidBrush(Color.White))
        {
            g.DrawString(week.ToString("D2", CultureInfo.InvariantCulture), font, brush, (size > 16 ? -fontSize * 0.12f : -fontSize * 0.07f), (size > 16 ? fontSize * 0.2f : fontSize * 0.08f));
        }
    }

    public static bool StartWithWindows
    {
        get
        {
            try
            {
                return Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", Application.ProductName, null) != null;
            }
            catch { return false; }
        }
        set
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (value)
                    {
                        key.SetValue(Application.ProductName, "\"" + Application.ExecutablePath + "\"");
                    }
                    else
                    {
                        key.DeleteValue(Application.ProductName, false);
                    }
                }
            }
            catch { }
        }
    }
}
