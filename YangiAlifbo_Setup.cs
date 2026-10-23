using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;

namespace YangiAlifboInstaller {
    public class InstallerForm : Form {
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblInfo;
        private Button btnInstall;
        private Button btnUninstall;
        private Label lblStatus;
        private ProgressBar progressBar;

        public InstallerForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "O'zbekcha Yangi Alifbo Klaviaturasi - O'rnatuvchi";
            this.Size = new Size(540, 420);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);

            // Title
            lblTitle = new Label();
            lblTitle.Text = "O'zbekcha Yangi Alifbo Klaviaturasi";
            lblTitle.Font = new Font("Segoe UI", 16f, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(15, 23, 42);
            lblTitle.Location = new Point(25, 25);
            lblTitle.Size = new Size(480, 35);

            // Subtitle
            lblSubtitle = new Label();
            lblSubtitle.Text = "Windows tizimining til qatoriga yangi alifboni (Ö, Ğ, Ş, Ç) qo'shish";
            lblSubtitle.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            lblSubtitle.ForeColor = Color.FromArgb(71, 85, 105);
            lblSubtitle.Location = new Point(27, 62);
            lblSubtitle.Size = new Size(480, 25);

            // Key info card
            Panel card = new Panel();
            card.Location = new Point(25, 100);
            card.Size = new Size(475, 110);
            card.BackColor = Color.FromArgb(241, 245, 249);
            card.BorderStyle = BorderStyle.FixedSingle;

            lblInfo = new Label();
            lblInfo.Text = "✨ Yangi harflar qulay tugmalarga biriktirilgan:\n" +
                           "  • [ tugmasida: ö va Ö    |  ; tugmasida: ş va Ş\n" +
                           "  • ] tugmasida: ğ va Ğ    |  ' tugmasida: ç va Ç\n" +
                           "  • \\ tugmasida: ; va :    |  ` tugmasida: ' va \"";
            lblInfo.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            lblInfo.ForeColor = Color.FromArgb(30, 41, 59);
            lblInfo.Location = new Point(12, 10);
            lblInfo.Size = new Size(450, 88);
            card.Controls.Add(lblInfo);

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Location = new Point(25, 225);
            progressBar.Size = new Size(475, 18);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;

            // Status label
            lblStatus = new Label();
            lblStatus.Text = "Klaviaturani o'rnatish uchun quyidagi tugmani bosing:";
            lblStatus.Location = new Point(25, 255);
            lblStatus.Size = new Size(475, 25);
            lblStatus.ForeColor = Color.FromArgb(100, 116, 139);

            // Install Button
            btnInstall = new Button();
            btnInstall.Text = "📥  Klaviaturani O'rnatish (1-klik)";
            btnInstall.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            btnInstall.BackColor = Color.FromArgb(16, 185, 129);
            btnInstall.ForeColor = Color.White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Location = new Point(25, 285);
            btnInstall.Size = new Size(320, 50);
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
            btnUninstall.Location = new Point(360, 285);
            btnUninstall.Size = new Size(140, 50);
            btnUninstall.Cursor = Cursors.Hand;
            btnUninstall.Click += BtnUninstall_Click;

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblSubtitle);
            this.Controls.Add(card);
            this.Controls.Add(progressBar);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnInstall);
            this.Controls.Add(btnUninstall);
        }

        private void BtnInstall_Click(object sender, EventArgs e) {
            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            progressBar.Visible = true;
            lblStatus.Text = "O'rnatilmoqda, iltimos kuting...";
            lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
            Application.DoEvents();

            try {
                DoInstall();
                progressBar.Visible = false;
                lblStatus.Text = "✅ Muvaffaqiyatli o'rnatildi!";
                MessageBox.Show(
                    "O'zbekcha yangi alifbo klaviaturasi tizimga muvaffaqiyatli o'rnatildi!\n\n" +
                    "Foydalanish uchun: Win + Bo'sh joy (Space) tugmalarini bosing va 'Uzbek - new' klaviaturasini tanlang.",
                    "O'rnatish yakunlandi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
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
                "O'zbekcha yangi alifbo klaviaturasini tizimdan o'chirmoqchimisiz?",
                "Tasdiqlash",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (res != DialogResult.Yes) return;

            btnInstall.Enabled = false;
            btnUninstall.Enabled = false;
            progressBar.Visible = true;
            lblStatus.Text = "Tizimdan o'chirilmoqda...";
            Application.DoEvents();

            try {
                DoUninstall();
                progressBar.Visible = false;
                lblStatus.Text = "Klaviaturani o'chirish yakunlandi.";
                MessageBox.Show("Klaviaturani o'chirish yakunlandi.", "Ma'lumot", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public static void DoInstall() {
            string tempDir = Path.Combine(Path.GetTempPath(), "YangiAlifbo_Setup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                ExtractResourceZip(tempDir);
                bool is64 = Environment.Is64BitOperatingSystem;
                string msiName = is64 ? "uz-new_amd64.msi" : "uz-new_i386.msi";
                string msiPath = Path.Combine(tempDir, msiName);

                if (!File.Exists(msiPath)) {
                    // Try looking inside subfolder
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
                    throw new Exception("MSI o'rnatish kodi: " + p.ExitCode);
                }
            } finally {
                try {
                    Directory.Delete(tempDir, true);
                } catch {}
            }
        }

        public static void DoUninstall() {
            string tempDir = Path.Combine(Path.GetTempPath(), "YangiAlifbo_Setup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                ExtractResourceZip(tempDir);
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
            } finally {
                try {
                    Directory.Delete(tempDir, true);
                } catch {}
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
                    DoInstall();
                    return;
                } else if (arg == "/uninstall" || arg == "-u") {
                    DoUninstall();
                    return;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
        }
    }
}
