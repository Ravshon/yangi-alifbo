using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;

namespace YangiAlifboInstaller {
    public class InstallerForm : Form {
        private Label lblTitle;
        private Label lblSubtitle;
        private CheckBox chkKeyboard;
        private CheckBox chkWordAddin;
        private Button btnInstall;
        private Button btnUninstall;
        private Label lblStatus;
        private ProgressBar progressBar;

        public InstallerForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "O'zbekcha Yangi Alifbo - O'rnatuvchi Dastur";
            this.Size = new Size(560, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);

            // Title
            lblTitle = new Label();
            lblTitle.Text = "O'zbekcha Yangi Alifbo To'plami";
            lblTitle.Font = new Font("Segoe UI", 16f, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(15, 23, 42);
            lblTitle.Location = new Point(25, 22);
            lblTitle.Size = new Size(500, 35);

            // Subtitle
            lblSubtitle = new Label();
            lblSubtitle.Text = "Windows klaviaturasi va Microsoft Word uchun 1-klikli to'liq yechim";
            lblSubtitle.Font = new Font("Segoe UI", 9.8f, FontStyle.Regular);
            lblSubtitle.ForeColor = Color.FromArgb(71, 85, 105);
            lblSubtitle.Location = new Point(27, 58);
            lblSubtitle.Size = new Size(500, 25);

            // Options Group Box
            GroupBox grp = new GroupBox();
            grp.Text = " O'rnatiladigan komponentlarni tanlang: ";
            grp.Location = new Point(25, 95);
            grp.Size = new Size(495, 180);
            grp.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            grp.ForeColor = Color.FromArgb(30, 41, 59);

            chkKeyboard = new CheckBox();
            chkKeyboard.Text = "1. Windows Klaviatura Drayveri (Ö, Ğ, Ş, Ç)\n   Standart QWERTY klaviaturada bitta bosishda yangi harflarni yozish imkoniyati.";
            chkKeyboard.Font = new Font("Segoe UI", 9.2f, FontStyle.Regular);
            chkKeyboard.ForeColor = Color.FromArgb(15, 23, 42);
            chkKeyboard.Location = new Point(20, 28);
            chkKeyboard.Size = new Size(460, 50);
            chkKeyboard.Checked = true;

            chkWordAddin = new CheckBox();
            chkWordAddin.Text = "2. Microsoft Word Qo'shimchasi (Add-in)\n   Word'da 'Вид'dan keyin 'Yangi Alifbo' lentasi: Kirill-Lotin va Yangi-Eski alifboga 1-klikda o'tkazish.";
            chkWordAddin.Font = new Font("Segoe UI", 9.2f, FontStyle.Regular);
            chkWordAddin.ForeColor = Color.FromArgb(15, 23, 42);
            chkWordAddin.Location = new Point(20, 95);
            chkWordAddin.Size = new Size(460, 55);
            chkWordAddin.Checked = true;

            grp.Controls.Add(chkKeyboard);
            grp.Controls.Add(chkWordAddin);

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Location = new Point(25, 290);
            progressBar.Size = new Size(495, 16);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;

            // Status label
            lblStatus = new Label();
            lblStatus.Text = "O'rnatish uchun quyidagi tugmani bosing:";
            lblStatus.Location = new Point(25, 316);
            lblStatus.Size = new Size(495, 25);
            lblStatus.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            lblStatus.ForeColor = Color.FromArgb(100, 116, 139);

            // Install Button
            btnInstall = new Button();
            btnInstall.Text = "🚀  O'rnatish (1-klik)";
            btnInstall.Font = new Font("Segoe UI", 11.5f, FontStyle.Bold);
            btnInstall.BackColor = Color.FromArgb(16, 185, 129);
            btnInstall.ForeColor = Color.White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Location = new Point(25, 348);
            btnInstall.Size = new Size(340, 52);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Click += BtnInstall_Click;

            // Uninstall Button
            btnUninstall = new Button();
            btnUninstall.Text = "🗑️  O'chirish";
            btnUninstall.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            btnUninstall.BackColor = Color.FromArgb(226, 232, 240);
            btnUninstall.ForeColor = Color.FromArgb(71, 85, 105);
            btnUninstall.FlatStyle = FlatStyle.Flat;
            btnUninstall.FlatAppearance.BorderSize = 0;
            btnUninstall.Location = new Point(380, 348);
            btnUninstall.Size = new Size(140, 52);
            btnUninstall.Cursor = Cursors.Hand;
            btnUninstall.Click += BtnUninstall_Click;

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblSubtitle);
            this.Controls.Add(grp);
            this.Controls.Add(progressBar);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnInstall);
            this.Controls.Add(btnUninstall);
        }

        private void BtnInstall_Click(object sender, EventArgs e) {
            if (!chkKeyboard.Checked && !chkWordAddin.Checked) {
                MessageBox.Show("Iltimos, kamida bitta komponentni tanlang.", "Ogohlantirish", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Check if Word is running before proceeding
            if (chkWordAddin.Checked && IsWordRunning()) {
                DialogResult dr = MessageBox.Show(
                    "Microsoft Word dasturi hozirda ochiq turibdi!\n\n" +
                    "Qo'shimchani to'liq o'rnatish uchun Word dasturini yopish kerak.\n\n" +
                    "Word dasturini hozir avtomatik yopaylikmi? (Hujjatingiz saqlanganligiga ishonch hosil qiling)",
                    "Microsoft Word ochiq",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question
                );

                if (dr == DialogResult.Yes) {
                    CloseWordProcesses();
                } else if (dr == DialogResult.No) {
                    MessageBox.Show("Iltimos, Word hujjatlaringizni saqlab, dasturni yoping va qaytadan 'O'rnatish' tugmasini bosing.", "Ma'lumot", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                } else {
                    return;
                }
            }

            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            progressBar.Visible = true;
            lblStatus.Text = "O'rnatilmoqda, iltimos kuting...";
            lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
            Application.DoEvents();

            try {
                DoInstall(chkKeyboard.Checked, chkWordAddin.Checked);
                progressBar.Visible = false;
                lblStatus.Text = "✅ Muvaffaqiyatli o'rnatildi!";
                
                string msg = "Tanlangan komponentlar muvaffaqiyatli o'rnatildi!\n\n";
                if (chkKeyboard.Checked) {
                    msg += "• Klaviatura: Win + Bo'sh joy (Space) tugmasi orqali 'Uzbek - new' ni tanlang.\n";
                }
                if (chkWordAddin.Checked) {
                    msg += "• Microsoft Word: Dasturni ochsangiz, 'Вид' lentasidan keyin alohida 'Yangi Alifbo' bo'limi paydo bo'ldi.\n";
                }
                MessageBox.Show(msg, "O'rnatish yakunlandi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                progressBar.Visible = false;
                lblStatus.Text = "Xatolik yuz berdi.";
                lblStatus.ForeColor = Color.Red;
                MessageBox.Show("Xatolik: " + ex.Message, "Xatolik", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                btnInstall.Enabled = true;
                btnUninstall.Enabled = true;
            }
        }

        private void BtnUninstall_Click(object sender, EventArgs e) {
            DialogResult res = MessageBox.Show(
                "Yangi alifbo komponentlarini tizimdan o'chirmoqchimisiz?",
                "Tasdiqlash",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (res != DialogResult.Yes) return;

            if (chkWordAddin.Checked && IsWordRunning()) {
                DialogResult dr = MessageBox.Show(
                    "Microsoft Word dasturi hozirda ochiq turibdi!\nWord dasturini avtomatik yopaylikmi?",
                    "Word ochiq",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (dr == DialogResult.Yes) CloseWordProcesses();
                else return;
            }

            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            progressBar.Visible = true;
            lblStatus.Text = "Tizimdan o'chirilmoqda...";
            Application.DoEvents();

            try {
                DoUninstall(chkKeyboard.Checked, chkWordAddin.Checked);
                progressBar.Visible = false;
                lblStatus.Text = "O'chirish yakunlandi.";
                MessageBox.Show("Tanlangan komponentlar tizimdan o'chirildi.", "Ma'lumot", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                progressBar.Visible = false;
                lblStatus.Text = "O'chirishda xatolik yuz berdi.";
                lblStatus.ForeColor = Color.Red;
                MessageBox.Show("Xatolik: " + ex.Message, "Xatolik", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                btnInstall.Enabled = true;
                btnUninstall.Enabled = true;
            }
        }

        private static bool IsWordRunning() {
            Process[] procs = Process.GetProcessesByName("winword");
            return procs != null && procs.Length > 0;
        }

        private static void CloseWordProcesses() {
            Process[] procs = Process.GetProcessesByName("winword");
            foreach (Process p in procs) {
                try {
                    p.CloseMainWindow();
                } catch {}
            }
            Thread.Sleep(1000);
            procs = Process.GetProcessesByName("winword");
            foreach (Process p in procs) {
                try {
                    p.Kill();
                } catch {}
            }
            Thread.Sleep(500);
        }

        public static void DoInstall(bool installKeyboard, bool installWord) {
            string tempDir = Path.Combine(Path.GetTempPath(), "YangiAlifbo_Setup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                ExtractResourceZip(tempDir);

                // 1. Install Keyboard
                if (installKeyboard) {
                    bool is64 = Environment.Is64BitOperatingSystem;
                    string msiName = is64 ? "uz-new_amd64.msi" : "uz-new_i386.msi";
                    string msiPath = Path.Combine(tempDir, msiName);

                    if (!File.Exists(msiPath)) {
                        string sub = Path.Combine(tempDir, "YangiAlifbo_Klaviatura", msiName);
                        if (File.Exists(sub)) msiPath = sub;
                    }

                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = "msiexec.exe";
                    psi.Arguments = "/i \"" + msiPath + "\" /qb-";
                    psi.UseShellExecute = true;
                    psi.Verb = "runas";

                    Process p = Process.Start(psi);
                    p.WaitForExit();
                    if (p.ExitCode != 0 && p.ExitCode != 3010) {
                        throw new Exception("Klaviatura MSI o'rnatish kodi: " + p.ExitCode);
                    }
                }

                // 2. Install Word Addin
                if (installWord) {
                    string dotmName = "YangiAlifbo.dotm";
                    string dotmSrc = Path.Combine(tempDir, dotmName);
                    if (!File.Exists(dotmSrc)) {
                        dotmSrc = Path.Combine(tempDir, "YangiAlifbo_Klaviatura", dotmName);
                    }

                    if (File.Exists(dotmSrc)) {
                        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        string wordStartup = Path.Combine(appData, "Microsoft", "Word", "STARTUP");
                        Directory.CreateDirectory(wordStartup);
                        string dotmDest = Path.Combine(wordStartup, dotmName);

                        // Try copying with retries in case lock is releasing
                        int retries = 3;
                        while (retries > 0) {
                            try {
                                File.Copy(dotmSrc, dotmDest, true);
                                break;
                            } catch (IOException) {
                                retries--;
                                if (retries == 0) {
                                    throw new Exception("Microsoft Word dasturi ochiq turgani sababli fayl qulflangan. Iltimos, Word dasturini to'liq yopib qaytadan urinib ko'ring.");
                                }
                                Thread.Sleep(1000);
                            }
                        }
                    }
                }
            } finally {
                try { Directory.Delete(tempDir, true); } catch {}
            }
        }

        public static void DoUninstall(bool uninstallKeyboard, bool uninstallWord) {
            string tempDir = Path.Combine(Path.GetTempPath(), "YangiAlifbo_Setup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                ExtractResourceZip(tempDir);

                if (uninstallKeyboard) {
                    bool is64 = Environment.Is64BitOperatingSystem;
                    string msiName = is64 ? "uz-new_amd64.msi" : "uz-new_i386.msi";
                    string msiPath = Path.Combine(tempDir, msiName);

                    if (!File.Exists(msiPath)) {
                        string sub = Path.Combine(tempDir, "YangiAlifbo_Klaviatura", msiName);
                        if (File.Exists(sub)) msiPath = sub;
                    }

                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = "msiexec.exe";
                    psi.Arguments = "/x \"" + msiPath + "\" /qb-";
                    psi.UseShellExecute = true;
                    psi.Verb = "runas";

                    Process p = Process.Start(psi);
                    p.WaitForExit();
                }

                if (uninstallWord) {
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string dotmDest = Path.Combine(appData, "Microsoft", "Word", "STARTUP", "YangiAlifbo.dotm");
                    if (File.Exists(dotmDest)) {
                        int retries = 3;
                        while (retries > 0) {
                            try {
                                File.Delete(dotmDest);
                                break;
                            } catch (IOException) {
                                retries--;
                                if (retries == 0) {
                                    throw new Exception("Word dasturi ochiq turgani sababli faylni o'chirib bo'lmadi. Iltimos, Word dasturini yoping.");
                                }
                                Thread.Sleep(1000);
                            }
                        }
                    }
                }
            } finally {
                try { Directory.Delete(tempDir, true); } catch {}
            }
        }

        private static void ExtractResourceZip(string targetDir) {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream s = asm.GetManifestResourceStream("payload.zip")) {
                if (s == null) throw new Exception("Ichki resurs topilmadi.");
                string zipPath = Path.Combine(targetDir, "payload.zip");
                using (FileStream fs = new FileStream(zipPath, FileMode.Create, FileAccess.Write)) {
                    s.CopyTo(fs);
                }
                ZipFile.ExtractToDirectory(zipPath, targetDir);
                File.Delete(zipPath);
            }
        }

        [STAThread]
        static void Main(string[] args) {
            if (args.Length > 0) {
                string arg = args[0].ToLowerInvariant();
                if (arg == "/install" || arg == "/silent" || arg == "-i") {
                    DoInstall(true, true);
                    return;
                } else if (arg == "/uninstall" || arg == "-u") {
                    DoUninstall(true, true);
                    return;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
        }
    }
}
