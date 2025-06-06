using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
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
    private static readonly System.Threading.Mutex Mutex = new System.Threading.Mutex(true, "A2F14B3D-7C9E-489F-8A76-3E7D925F146C");

    [STAThread]
    static void Main()
    {
        if (!Mutex.WaitOne(TimeSpan.Zero, true))
        {
            return;
        }

        Application.SetCompatibleTextRenderingDefault(false);
        Application.EnableVisualStyles();
        SetProcessDpiAwarenessContext(new IntPtr(-4));

        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer() { Interval = 10000, Enabled = true };
        timer.Tick += new EventHandler(TimerTick);

        currentWeek = GetWeek();
        weekIcon = GetIcon(currentWeek);
        icon = new NotifyIcon()
        {
            Icon = weekIcon,
            Visible = true,
            ContextMenu = new ContextMenu(new MenuItem[]
            {
                new MenuItem("About", new EventHandler(AboutClick)),
                new MenuItem("Open web site...", new EventHandler(OpenWebsiteClick)),
                new MenuItem("Save Icon...", new EventHandler(SaveIconClick)),
                new MenuItem("Start with Windows", new EventHandler(StartWithWindowsClick)) { Checked = StartWithWindows },
                new MenuItem("Exit", new EventHandler(ExitClick))
            })
        };

        UpdateIcon();
        Application.Run(new ApplicationContext());
    }

    private static void TimerTick(object sender, EventArgs e)
    {
        UpdateIcon();
    }

    private static void AboutClick(object sender, EventArgs e)
    {
        MessageBox.Show(string.Format("{0} - v{1}", Application.ProductName, Application.ProductVersion)); 
    }

    private static void OpenWebsiteClick(object sender, EventArgs e)
    {
        Process.Start("https://voltura.github.io/WeekNumberLite2Plus"); 
    }

    private static void SaveIconClick(object sender, EventArgs e)
    {
        using (var fs = File.Create(Path.Combine(Environment.GetFolderPath((Environment.SpecialFolder)0x10), "WeekIcon.ico")))
        {
            weekIcon.Save(fs);
        }
    
        MessageBox.Show("Icon saved to desktop.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static void StartWithWindowsClick(object sender, EventArgs e)
    {
        MenuItem menuItem = (MenuItem)sender;
        StartWithWindows = !menuItem.Checked;
        menuItem.Checked = StartWithWindows;
    }

    private static void ExitClick(object sender, EventArgs e)
    {
        if (icon != null)
        {
            icon.Visible = false;
            DestroyIcon(icon.Icon.Handle);
        }

        if (weekIcon != null)
        {
            DestroyIcon(weekIcon.Handle);
        }

        Application.Exit();
    }

    private static int GetWeek()
    {
        DateTime now = DateTime.Now;
        int year = now.Year;
        int week = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        int p1 = (year + year / 4 - year / 100 + year / 400) % 7;
        int p2 = (year - 1 + (year - 1) / 4 - (year - 1) / 100 + (year - 1) / 400) % 7;

        return (week == 53 && !(p1 == 4 || p2 == 3)) ? 1 : week;
    }

    private static void UpdateIcon()
    {
        try
        {
            int week = GetWeek();
            DateTime now = DateTime.Now;

            icon.Text = string.Format("{0} {1}\r\n{2}-{3:D2}-{4:D2}", CultureInfo.CurrentUICulture.LCID == 1053 ? "Vecka" : "Week",
                week, now.Year, now.Month, now.Day);

            if (week != currentWeek)
            {
                Icon prevIcon = icon.Icon;

                weekIcon = GetIcon(week);
                icon.Icon = weekIcon;

                currentWeek = week;

                if (prevIcon != null)
                {
                    DestroyIcon(prevIcon.Handle);
                    prevIcon.Dispose();
                }
            }
        }
        catch { }
    }

    private static readonly int[] sizes = { 16, 32, 48 };

    internal static Icon GetIcon(int week)
    {
        using (MemoryStream iconStream = new MemoryStream())
        using (BinaryWriter writer = new BinaryWriter(iconStream))
        {
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write((ushort)3);

            int imageOffset = 54;
            byte[][] images = new byte[3][];

            using (MemoryStream ms = new MemoryStream())
            {
                for (int i = 0, size; i < 3; i++)
                {
                    size = sizes[i];
                    ms.SetLength(0);

                    using (Bitmap bmp = new Bitmap(size, size))
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                        DrawBackground(g, size);
                        DrawWeekText(g, size, week);

                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        images[i] = ms.ToArray();
                    }
                }
            }

            for (int i = 0; i < 3; i++)
            {
                writer.Write((byte)sizes[i]);
                writer.Write((byte)sizes[i]);
                writer.Write((byte)0);
                writer.Write((byte)0);
                writer.Write((ushort)1);
                writer.Write((ushort)32);
                writer.Write(images[i].Length);
                writer.Write((uint)imageOffset);
                imageOffset += images[i].Length;
            }

            for (int i = 0; i < images.Length; i++)
            {
                writer.Write(images[i]);
            }

            iconStream.Position = 0;

            return new Icon(iconStream);
        }
    }

    private static void DrawBackground(Graphics g, int size)
    {
        int inset = size >> 5;

        using (SolidBrush bg = new SolidBrush(Color.Black))
        using (SolidBrush fg = new SolidBrush(Color.LightGray))
        using (Pen border = new Pen(Color.LightGray, inset << 1))
        {
            g.FillRectangle(bg, inset, inset, size - inset, size - inset);
            g.DrawRectangle(border, inset, inset, size - (inset << 1), size - (inset << 1));
            g.FillRectangle(fg, size >> 3, inset >> 1, inset * 3, inset * 5);
            g.FillRectangle(fg, size - (size >> 2), inset >> 1, inset * 3, inset * 5);
        }
    }

    private static void DrawWeekText(Graphics g, int size, int week)
    {
        int fontSize = (size * 25) >> 5;

        using (Font font = new Font(FontFamily.GenericMonospace, fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
        using (Brush brush = new SolidBrush(Color.White))
        {
            g.DrawString(week.ToString("D2"), font, brush, (size > 16) ? -fontSize * 0.12f : -fontSize * 0.07f, (size > 16) ? fontSize * 0.2f : fontSize * 0.08f);
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
            catch
            {
                return false;
            }
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
