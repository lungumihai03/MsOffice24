using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;


namespace MsOffice24
{
    public partial class Form1 : Form
    {
        // Definirea variabilelor
        string V2024 = "PerpetualVL2024";
        string V2021 = "PerpetualVL2021";
        string V2019 = "PerpetualVL2019";
        string V2016 = "PerpetualVL2016";
        string v24 = "ProPlus2024Volume";
        string v21 = "ProPlus2021Volume";
        string v19 = "ProPlus2019Volume";
        string v16 = "ProPlus2016Volume";
        string v;
        string version;
        string x64 = "64";
        string x86 = "32";
        string arch;
        string lang;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.Text = "Select an Office Application";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Setarea variabilelor pe baza selecției utilizatorului
            if (comboBox1.SelectedIndex == 0)
            {
                version = V2024;
                v = v24;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                version = V2021;
                v = v21;
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                version = V2019;
                v = v19;
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                version = V2016;
                v = v16;
            }

            arch = A64.Checked ? x64 : x86;

            lang = en.Checked ? en.Text :
                   fr.Checked ? fr.Text :
                   de.Checked ? de.Text :
                   es.Checked ? es.Text :
                   it.Checked ? it.Text :
                   ro.Checked ? ro.Text :
                   ru.Checked ? ru.Text :
                   cn.Checked ? cn.Text :
                   jp.Checked ? jp.Text :
                   pt.Checked ? pt.Text :
                   string.Empty;

            
            string filePath = "res/config.xml";
            string xmlContent = $@"<Configuration>
  <Add OfficeClientEdition=""{arch}"" Channel=""{version}"">
    <Product ID=""{v}"">
      <Language ID=""{lang}"" />
    </Product>
  </Add>
  <!--  <Updates Enabled=""TRUE"" Channel=""Current"" /> -->

  <!--  <Display Level=""None"" AcceptEULA=""TRUE"" />  -->

  <!--  <Property Name=""AUTOACTIVATE"" Value=""1"" />  -->

</Configuration>
";

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.Write(xmlContent);
            }

            if (checkBox1.Checked)
            {
                ProcessStartInfo psStartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "reg add \"HKCU\\Software\\Microsoft\\Office\\16.0\\Common\\ExperimentConfigs\\Ecs\" /v \"CountryCode\" /t REG_SZ /d \"std::wstring|US\" /f",
                    Verb = "runas", // Asta va cere permisiuni de administrator
                    UseShellExecute = true
                };
                Process.Start(psStartInfo);
                MessageBox.Show("CountryCode set to US", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Console.WriteLine($"Fișierul {filePath} a fost creat cu succes.");

            // Rulează procesele în mod secvențial
            RunProcessSequentially("res/setup.exe", "/download res/config.xml", "/configure res/config.xml");
            if(checkBox2.Checked)
            {
                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    FileName = "powershell.exe",
                    Arguments = "-Command \"irm https://get.activated.win | iex\"",
                    Verb = "runas", // Execută cu privilegii de administrator
                    UseShellExecute = true
                };
                Process process = Process.Start(startInfo);

                // Așteptăm să se deschidă fereastra CMD
                process.WaitForExit();

            }
        }

        private void RunProcessSequentially(string filePath, string arguments1, string arguments2)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    // Rulează prima comandă
                    RunProcess(filePath, arguments1);

                    // Rulează a doua comandă
                    RunProcess(filePath, arguments2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Eroare: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Fișierul executabil nu a fost găsit.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RunProcess(string filePath, string arguments)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = filePath;
                process.StartInfo.Arguments = arguments;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;

                process.Start();
                process.WaitForExit(); // Așteaptă finalizarea procesului

                // Citirea ieșirii și erorilor (opțional)
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                Console.WriteLine("Ieșire: " + output);
                Console.WriteLine("Eroare: " + error);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }
    }
}
