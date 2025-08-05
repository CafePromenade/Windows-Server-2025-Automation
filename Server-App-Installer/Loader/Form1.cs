using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Loader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            new WebClient().DownloadFile("https://raw.githubusercontent.com/CafePromenade/Windows-Server-2025-Automation/refs/heads/main/Server-App-Installer/Server-App-Installer/bin/Debug/Server-App-Installer.exe", Environment.GetEnvironmentVariable("TEMP") + "\\ServerInstaller.exe");
            Process.Start(Environment.GetEnvironmentVariable("TEMP") + "\\ServerInstaller.exe");
            Close();
        }
    }
}
