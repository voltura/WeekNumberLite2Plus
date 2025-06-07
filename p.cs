using Microsoft.Win32;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Windows.Forms;

static class P
{
    [DllImport("user32.dll")]
    static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);

    [DllImport("user32.dll")]
    static extern bool DestroyIcon(IntPtr handle);

    static NotifyIcon icon;
    static Icon weekIcon;
    static int currentWeek;
    static readonly System.Threading.Mutex Mutex = new System.Threading.Mutex(true, "A2F14B3D-7C9E-489F-8A76-3E7D925F146C");
    static readonly string TranslatedWeekText = CultureInfo.CurrentUICulture.LCID == 1053 ? "Vecka" : "Week";

    [STAThread]
    static void Main()
    {
        if (!Mutex.WaitOne(TimeSpan.Zero, true)) return;

        GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
        GCSettings.LatencyMode = GCLatencyMode.Batch;
        Application.SetCompatibleTextRenderingDefault(false);
        Application.EnableVisualStyles();
        SetProcessDpiAwarenessContext(new IntPtr(-4));

        (new Timer() { Interval = 1800000, Enabled = true }).Tick += UpdateIconTimerTick;

        icon = new NotifyIcon()
        {
            Icon = weekIcon = GetIcon(currentWeek = GetWeek()),
            Text = string.Format("{0} {1}\r\n{2:yyyy-MM-dd}", TranslatedWeekText, currentWeek, DateTime.Now),
            Visible = true,
            ContextMenu = new ContextMenu(new MenuItem[]
            {
                new MenuItem("About", AboutClick),
                new MenuItem("Open web site", OpenWebsiteClick),
                new MenuItem("Save Icon", SaveIconClick),
                new MenuItem("Start with Windows", StartWithWindowsClick) { Checked = StartWithWindows },
                new MenuItem("Exit", ExitClick)
            })
        };

        Application.Run();
    }

    static void UpdateIconTimerTick(object sender, EventArgs e) { UpdateIcon(); }

    static void AboutClick(object sender, EventArgs e)
    {
        MessageBox.Show(string.Format("{0} - v{1}", Application.ProductName, Application.ProductVersion), "About");
    }

    static void OpenWebsiteClick(object sender, EventArgs e)
    {
        System.Diagnostics.Process.Start("https://voltura.github.io/WeekNumberLite2Plus");
    }

    static void SaveIconClick(object sender, EventArgs e)
    {
        using (var fs = File.Create(Path.Combine(Environment.GetFolderPath((Environment.SpecialFolder)0x10), "WeekIcon.ico")))
            weekIcon.Save(fs);
        MessageBox.Show("Icon saved to desktop.", "Saved");
    }

    static void StartWithWindowsClick(object sender, EventArgs e)
    {
        MenuItem menuItem = (MenuItem)sender;
        menuItem.Checked = StartWithWindows = !menuItem.Checked;
    }

    static void ExitClick(object sender, EventArgs e)
    {
        if (icon != null)
        {
            icon.Visible = false;
            DestroyIcon(icon.Icon.Handle);
        }

        Application.Exit();
    }

    static int GetWeek()
    {
        DateTime now = DateTime.Now;
        int year = now.Year, week = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday),
            p1 = (year + year / 4 - year / 100 + year / 400) % 7, p2 = (year - 1 + (year - 1) / 4 - (year - 1) / 100 + (year - 1) / 400) % 7;

        return (week == 53 && !(p1 == 4 || p2 == 3)) ? 1 : week;
    }

    static void UpdateIcon()
    {
        int week = GetWeek();

        icon.Text = string.Format("{0} {1}\r\n{2:yyyy-MM-dd}", TranslatedWeekText, week, DateTime.Now);

        if (week != currentWeek)
        {
            Icon prevIcon = icon.Icon;

            icon.Icon = weekIcon = GetIcon(week);
            currentWeek = week;

            if (prevIcon != null)
            {
                DestroyIcon(prevIcon.Handle);
                prevIcon.Dispose();
            }
        }
    }

    static Icon GetIcon(int week)
    {
        using (MemoryStream iconStream = new MemoryStream())
        using (BinaryWriter writer = new BinaryWriter(iconStream))
        {
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write((ushort)2);

            int imageOffset = 38;
            byte[][] images = new byte[2][];

            using (MemoryStream ms = new MemoryStream())
            {
                for (int i = 0, size; i < 2; i++)
                {
                    size = (i == 0) ? 16 : 48;
                    ms.SetLength(0);

                    using (Bitmap bmp = new Bitmap(size, size))
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        int inset = Math.Max(1, size / 16);
                        int fontSize = ((size * 25) >> 5) - 1;

                        using (SolidBrush bg = new SolidBrush(Color.Black))
                        using (SolidBrush fg = new SolidBrush(Color.White))
                        using (Pen border = new Pen(Color.LightGray, size == 16 ? 0.1f : 0.2f))
                        using (Font font = new Font(FontFamily.GenericMonospace, fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                        {
                            int tabWidth = Math.Max(1, size / 10);
                            int tabHeight = Math.Max(1, size / 4) - 2;
                            int tabTop = inset >> 1;
                            int tabSpacing = Math.Max(2, (size - 2 * tabWidth) / 5);

                            g.FillRectangle(bg, inset, inset, size - inset * 2, size - inset * 2);
                            g.DrawRectangle(border, inset, inset, size - inset * 2 - 1, size - inset * 2 - 1);
                            g.FillRectangle(fg, tabSpacing + 1, tabTop, tabWidth, tabHeight);
                            g.FillRectangle(fg, size - tabSpacing - tabWidth - 1, tabTop, tabWidth, tabHeight);
                            g.DrawString(week.ToString("D2"), font, fg, (size > 16) ? -fontSize * 0.10f : -fontSize * 0.07f, (size > 16) ? fontSize * 0.16f : fontSize * 0.08f);
                        }

                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        images[i] = ms.ToArray();
                    }
                }
            }

            for (int i = 0, size; i < 2; i++)
            {
                size = (i == 0) ? 16 : 48;
                writer.Write((byte)size);
                writer.Write((byte)size);
                writer.Write((byte)0);
                writer.Write((byte)0);
                writer.Write((ushort)1);
                writer.Write((ushort)32);
                writer.Write(images[i].Length);
                writer.Write((uint)imageOffset);
                imageOffset += images[i].Length;
            }

            for (int i = 0; i < images.Length; i++)
                writer.Write(images[i]);
            
            iconStream.Position = 0;

            return new Icon(iconStream);
        }
    }

    static bool StartWithWindows
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
                    if (value) key.SetValue(Application.ProductName, "\"" + Application.ExecutablePath + "\"");
                    else key.DeleteValue(Application.ProductName, false);
            }
            catch { }
        }
    }
}
