using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NetSpeedTray
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool createdNew = true;
            using (Mutex mutex = new Mutex(true, "NetSpeedTray", out createdNew))
            {
                if (createdNew)
                {
                    Application.SetHighDpiMode(HighDpiMode.SystemAware);
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    Application.Run(new MyCustomApplicationContext());
                }
                else
                {
                    MessageBox.Show("NetSpeedTray is already running and silently lives on the taskbar notification area.", "NetSpeedTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        public class MyCustomApplicationContext : ApplicationContext
        {
            static System.Timers.Timer t;

            private NotifyIcon trayIcon;
            private ContextMenuStrip contextMenuStrip;
            private ToolStripMenuItem toolStripSettings = new ToolStripMenuItem();
            private ToolStripMenuItem toolStripQuit = new ToolStripMenuItem();

            private Settings s = new Settings();

            private System.Windows.Forms.Timer doubleClickTimer = new System.Windows.Forms.Timer();

            public MyCustomApplicationContext()
            {
#if DEBUG
                s.Show();
#endif

                InitTimer();

                doubleClickTimer.Interval = 100;
                doubleClickTimer.Tick += new EventHandler(doubleClickTimer_Tick);

                toolStripSettings.Text = "Settings";
                toolStripSettings.Click += ToolStripOpen_Click;
                toolStripSettings.AutoSize = false;
                toolStripSettings.Size = new Size(120, 30);
                toolStripSettings.Margin = new Padding(0, 4, 0, 0);

                toolStripQuit.Text = "Quit";
                toolStripQuit.Click += ToolStripExit_Click;
                toolStripQuit.AutoSize = false;
                toolStripQuit.Size = new Size(120, 30);
                toolStripQuit.Margin = new Padding(0, 0, 0, 4);

                contextMenuStrip = new ContextMenuStrip()
                {
                    DropShadowEnabled = true,
                    ShowCheckMargin = false,
                    ShowImageMargin = false,
                    Size = new System.Drawing.Size(310, 170)
                };

                contextMenuStrip.Items.AddRange(new ToolStripItem[]
                {
                    toolStripSettings,
                    toolStripQuit
                });

                try
                {
                    CustomizeContextMenuBackground();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something went wrong, try again.",
                                            "NetSpeedTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                trayIcon = new NotifyIcon()
                {
                    ContextMenuStrip = contextMenuStrip,
                    Visible = true,
                    Text = "NetSpeedTray"
                };

                trayIcon.MouseDown += TrayIcon_MouseDown;
                trayIcon.MouseClick += TrayIcon_MouseClick;

                void InitTimer()
                {
                    t = new System.Timers.Timer();
                    t.AutoReset = false;
                    t.Elapsed += new System.Timers.ElapsedEventHandler(t_Elapsed);
                    t.Interval = 1000; // milliseconds
                    t.Enabled = true;
                    t.Start();
                }


                void t_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
                {
                    Debug.WriteLine(DateTime.Now.ToString("hh:mm tt"));

                    // Read netspeed
                    var adapter = NetworkInterface.GetAllNetworkInterfaces()[Properties.Settings.Default.Device];
                    var reads = Enumerable.Empty<double>();
                    var readsUp = Enumerable.Empty<double>();
                    var sw = new Stopwatch();
                    var lastBr = adapter.GetIPv4Statistics().BytesReceived;
                    var lastBs = adapter.GetIPv4Statistics().BytesSent;
                    for (var i = 0; i < 1000; i++)
                    {

                        sw.Restart();
                        Thread.Sleep(100);
                        var elapsed = sw.Elapsed.TotalSeconds;
                        var br = adapter.GetIPv4Statistics().BytesReceived;
                        var bs = adapter.GetIPv4Statistics().BytesSent;

                        var local = (br - lastBr) / elapsed;
                        var localUp = (bs - lastBs) / elapsed;
                        lastBr = br;
                        lastBs = bs;

                        // Keep last 20, ~2 seconds
                        reads = new[] { local }.Concat(reads).Take(20);
                        readsUp = new[] { localUp }.Concat(readsUp).Take(20);

                        if (i % 10 == 0)
                        { // ~1 second
                            var bSec = reads.Sum() / reads.Count();
                            var bSecUp = readsUp.Sum() / readsUp.Count();
                            var mbs = (bSec * 8) / 1024 / 1024;
                            var mbsUp = (bSecUp * 8) / 1024 / 1024;
                            Debug.WriteLine(mbs.ToString("0") + " Mbps");
                            CreateTextIcon(mbs.ToString("0"), mbsUp.ToString("0"));
                        }
                    }
                    t.Interval = 1000;
                    t.Start();
                }


                void CreateTextIcon(string speed, string upSpeed)
                {
                    Font fontToUse = new Font("Trebuchet MS", 12, FontStyle.Regular, GraphicsUnit.Pixel);
                    Brush brushToUse = new SolidBrush(GetAccentColor());
                    Bitmap bitmapText = new Bitmap(16, 16);
                    Graphics g = Graphics.FromImage(bitmapText);

                    IntPtr hIcon;

                    g.Clear(Color.Transparent);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
                    g.DrawString(speed, fontToUse, brushToUse, GetPadding(speed), 0);
                    hIcon = bitmapText.GetHicon();
                    trayIcon.Icon = Icon.FromHandle(hIcon);
                    trayIcon.Text = "Up: " + upSpeed + " Mbps \n" + "Down: " + speed + " Mbps";
                }
            }

            private float GetPadding(string speed)
            {
                return Convert.ToInt32(speed) <= 9 ? 4 : 0;
            }

            private Color GetAccentColor()
            {
                Color accent = Color.White;
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent"))
                {
                    if (key != null)
                    {
                        var keyValue = key.GetValue("AccentColorMenu");
                        if (keyValue == null)
                        {
                            accent = Color.Lime;
                        }
                        else
                        {
                            // Convert Hex to Color
                            accent = HexToColor(Convert.ToInt32(keyValue).ToString("X8"));
                        }
                    }
                    else
                    {
                        accent = Color.Lime;
                    }
                }

                return accent;
            }

            public static Color HexToColor(string hexString)
            {
                // Remove #
                if (hexString.IndexOf('#') != -1)
                {
                    hexString = hexString.Replace("#", "");
                }

                int a, r, g, b = 0;
                // Key value format: AABBGGRR
                // Use return as is
                // TODO: Calculate for contrast

                a = int.Parse(hexString.Substring(0, 2), NumberStyles.AllowHexSpecifier);
                r = int.Parse(hexString.Substring(2, 2), NumberStyles.AllowHexSpecifier);
                g = int.Parse(hexString.Substring(4, 2), NumberStyles.AllowHexSpecifier);
                b = int.Parse(hexString.Substring(6, 2), NumberStyles.AllowHexSpecifier);

                Debug.WriteLine($"{a} {r} {g} {b}");
                return Color.FromArgb(a, r, g, b);
            }

            private async void CustomizeContextMenuBackground()
            {
                var verticalPadding = 4;
                contextMenuStrip.Items[0].Font = new Font(this.contextMenuStrip.Items[0].Font, FontStyle.Bold);
                bool appsUseLight = await Task.Run(() => ReadRegistry());

                if (appsUseLight)
                {
                    contextMenuStrip.Renderer = new MyCustomRenderer { VerticalPadding = verticalPadding, HighlightColor = Color.White, ImageColor = Color.FromArgb(255, 238, 238, 238) };
                    contextMenuStrip.BackColor = Lighten(Color.White);
                    contextMenuStrip.ForeColor = Color.Black;
                }
                else
                {
                    contextMenuStrip.Renderer = new MyCustomRenderer { VerticalPadding = verticalPadding, HighlightColor = Color.Black, ImageColor = Color.FromArgb(255, 43, 43, 43) };
                    contextMenuStrip.BackColor = Lighten(Color.Black);
                    contextMenuStrip.ForeColor = Color.White;
                }

                contextMenuStrip.MinimumSize = new Size(120, 30);
                contextMenuStrip.AutoSize = false;
                contextMenuStrip.ShowImageMargin = false;
                contextMenuStrip.ShowCheckMargin = false;
            }
            private bool ReadRegistry()
            {
                bool isUsingLightTheme;
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (key == null)
                        isUsingLightTheme = true;
                    else
                    {
                        var keyValue = key.GetValue("AppsUseLightTheme");
                        isUsingLightTheme = keyValue != null ? keyValue.ToString() == "1" : true;
                    }
                }
                return isUsingLightTheme;
            }


            private Color Lighten(Color color)
            {
                int r;
                int g;
                int b;

                if (color.R == 0 && color.G == 0 && color.B == 0)
                {
                    r = color.R + 43;
                    g = color.G + 43;
                    b = color.B + 43;
                }
                else
                {
                    r = color.R - 17;
                    g = color.G - 17;
                    b = color.B - 17;
                }

                return Color.FromArgb(r, g, b);
            }

            private bool isFirstClick = true;
            private bool isDoubleClick = false;
            private int milliseconds = 0;

            private void TrayIcon_MouseDown(object sender, MouseEventArgs e)
            {
                // This is the first mouse click.
                if (e.Button == MouseButtons.Left)
                {
                    if (isFirstClick)
                    {
                        isFirstClick = false;

                        // Start the double click timer.
                        doubleClickTimer.Start();
                    }

                    // This is the second mouse click.
                    else
                    {
                        // Verify that the mouse click is within the double click
                        // rectangle and is within the system-defined double 
                        // click period.
                        if (milliseconds < SystemInformation.DoubleClickTime)
                        {
                            isDoubleClick = true;
                        }
                    }
                }
            }

            private void doubleClickTimer_Tick(object sender, EventArgs e)
            {
                try
                {
                    milliseconds += 50;

                    // The timer has reached the double click time limit.
                    if (milliseconds >= SystemInformation.DoubleClickTime)
                    {
                        doubleClickTimer.Stop();

                        if (isDoubleClick)
                        {
                            Debug.WriteLine("Doubleclick");
                            // Perform Double Click
                            s.Show();
                        }

                        // Allow the MouseDown event handler to process clicks again.
                        isFirstClick = true;
                        isDoubleClick = false;
                        milliseconds = 0;
                    }


                }
                catch (Exception) { }
            }

            private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right)
                {
                    try
                    {
                        CustomizeContextMenuBackground();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Something went wrong, try again.",
                                                "NetSpeedTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            private void ToolStripExit_Click(object sender, EventArgs e)
            {
                // Hide tray icon, otherwise it will remain shown until user mouses over it
                trayIcon.Visible = false;
                Environment.Exit(0);
                Application.Exit();
            }

            private void ToolStripOpen_Click(object sender, EventArgs e)
            {
                try
                {
                    s.Show();
                }
                catch (Exception)
                {
                }
            }

            void Exit(object sender, EventArgs e)
            {
                // Hide tray icon, otherwise it will remain shown until user mouses over it
                trayIcon.Visible = false;

                Application.Exit();
            }
        }

        public class MyColorTable : ProfessionalColorTable
        {
            public override Color ToolStripGradientBegin
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color ToolStripGradientEnd
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color MenuItemBorder
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color MenuItemSelected
            {
                get { return Color.WhiteSmoke; }
            }
            public override Color ToolStripDropDownBackground
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color ImageMarginGradientBegin
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color ImageMarginGradientMiddle
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
            public override Color ImageMarginGradientEnd
            {
                get { return Color.FromArgb(255, 43, 43, 43); }
            }
        }


        private class MyCustomRenderer : ToolStripProfessionalRenderer
        {
            public MyCustomRenderer() : base(new MyColorTable())
            {
            }

            public Color ImageColor { get; set; }
            public Color HighlightColor { get; set; }
            public int VerticalPadding { get; set; }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                if (null == e)
                { return; }
                e.TextFormat &= ~TextFormatFlags.HidePrefix;
                e.TextFormat |= TextFormatFlags.VerticalCenter;
                var rect = e.TextRectangle;
                rect.Offset(24, VerticalPadding);
                e.TextRectangle = rect;
                base.OnRenderItemText(e);
            }

            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs myMenu)
            {
                if (!myMenu.Item.Selected)
                    base.OnRenderMenuItemBackground(myMenu);
                else
                {
                    if (myMenu.Item.Enabled)
                    {
                        Rectangle menuRectangle = new Rectangle(Point.Empty, myMenu.Item.Size);
                        //Fill Color
                        myMenu.Graphics.FillRectangle(new SolidBrush(RenderHighlight(HighlightColor)), menuRectangle);
                        // Border Color
                        // myMenu.Graphics.DrawRectangle(Pens.Lime, 1, 0, menuRectangle.Width - 2, menuRectangle.Height - 1);
                    }
                    else
                    {
                        Rectangle menuRectangle = new Rectangle(Point.Empty, myMenu.Item.Size);
                        //Fill Color
                        myMenu.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(20, 128, 128, 128)), menuRectangle);
                    }

                }
            }

            private Color RenderHighlight(Color color)
            {
                int r;
                int g;
                int b;

                if (color.R == 0 && color.G == 0 && color.B == 0)
                {
                    r = color.R + 65;
                    g = color.G + 65;
                    b = color.B + 65;
                }
                else
                {
                    r = color.R;
                    g = color.G;
                    b = color.B;
                }

                return Color.FromArgb(r, g, b);
            }

            protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new Rectangle(e.ImageRectangle.Location, e.ImageRectangle.Size);
                r.Inflate(1, 1);
                e.Graphics.FillRectangle(new SolidBrush(ImageColor), r);
                //r.Inflate(-4, -4);
                e.Graphics.DrawLines(Pens.Gray, new Point[]
                {
                    new Point(r.Left + 4, 10), //2
                    new Point(r.Left - 2 + r.Width / 2,  r.Height / 2 + 4), //3
                    new Point(r.Right - 4, r.Top + 4)
                });
            }
        }

    }
}
