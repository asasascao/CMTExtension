using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Assemblies;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MTEInstaller
{
    public partial class Form1 : Form
    {
        string regasmpath = "regasm";
        string regasmpath64 = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\regasm";
        string regasmpath32 = "C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\regasm";
        string regpath = "";

        string dllpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory)+@"\MoveToExtension.dll";
        public Form1()
        {
            InitializeComponent();

            string version = ReadVersion();
            label_version.Text = "版本号:" + version;
            if(string.IsNullOrWhiteSpace(version))
            {
                btn_install.Enabled = true;
                btn_uninstall.Enabled = false;
            }
            else
            {
                btn_install.Enabled = false;
                btn_uninstall.Enabled = true;
            }

            InitRegAsm();
        }

        private static string ReadVersion()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\CaoMoveToExtension", true))
            {
                if (key == null) return "";
                object value = key.GetValue("Ver");
                return value==null?"":value.ToString();
            }
        }

        private static void WriteVersion(string versionNumber)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\CaoMoveToExtension", true);
            if (key == null) key = Registry.CurrentUser.CreateSubKey(@"Software\CaoMoveToExtension");
            key.SetValue("Ver", versionNumber, RegistryValueKind.String);
            key.Close();
        }

        private static void DeleteVersion()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\CaoMoveToExtension", true);
            if (key == null) key = Registry.CurrentUser.CreateSubKey(@"Software\CaoMoveToExtension");
            key.DeleteValue("Ver");
            key.Close();
            Registry.CurrentUser.DeleteSubKey(@"Software\CaoMoveToExtension");
        }

        public void InitRegAsm()
        {
            regpath = regasmpath;
            var is64bit = Environment.Is64BitOperatingSystem;
            if (!CheckRegAsmEffective(regpath))
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    regpath = regasmpath64;
                    if (!CheckRegAsmEffective(regpath))
                    {
                        regpath = "";
                        MessageBox.Show("请先安装.net framework 4.0运行库,再运行本程序");
                    }
                }
                else
                {
                    regpath = regasmpath32;
                    if (!CheckRegAsmEffective(regpath))
                    {
                        regpath = "";
                        MessageBox.Show("请先安装.net framework 4.0运行库,再运行本程序");
                    }
                }
            }
        }

        private void btn_install_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(regpath)) return;

            FileVersionInfo version = FileVersionInfo.GetVersionInfo(dllpath);
            string versionNumber = version.FileVersion;

            if (!Install(out string msg))
            {
                MessageBox.Show(msg);
                return;
            }
            RestartExplorer();

            label_version.Text = "版本号:"+versionNumber;
            WriteVersion(versionNumber);
            MessageBox.Show("安装完成");
            btn_install.Enabled = false;
            btn_uninstall.Enabled = true;
        }

        private void btn_uninstall_Click(object sender, EventArgs e)
        {
            if (!UnInstall(out string msg))
            {
                MessageBox.Show(msg);
                return;
            }
            RestartExplorer();

            DeleteVersion();
            MessageBox.Show("卸载完成");
            btn_install.Enabled = true;
            btn_uninstall.Enabled = false;
            label_version.Text = "版本号:";
        }

        public bool CheckRegAsmEffective(string filename)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = filename,
                    Arguments = "",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    return !output.Contains("不是内部或外部命令");
                }
            }
            catch
            {
                return false;
            }
        }
        
        public void RestartExplorer()
        {
            foreach (Process proc in Process.GetProcessesByName("explorer"))
            {
                proc.Kill();
            }
            System.Threading.Thread.Sleep(1000);
            Process.Start("explorer.exe");
        }

        public bool Install(out string msg)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = regpath,
                Arguments = $" /codebase  MoveToExtension.dll",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                msg = output;
                return output.Contains("成功注册");
            }
        }

        public bool UnInstall(out string msg)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = regpath,
                Arguments = $" /codebase /unregister MoveToExtension.dll",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                msg = output;
                return output.Contains("成功注销");
            }
        }

        private void btn_update_Click(object sender, EventArgs e)
        {
            string filePath = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "C:\\Users\\Administrator\\Desktop"; // 设置初始目录
                openFileDialog.Title = "选择一个文件"; // 设置对话框标题
                openFileDialog.Filter = "升级文件 (*.zip)|*.zip"; // 设置文件过滤器
                openFileDialog.RestoreDirectory = true; // 打开对话框后是否恢复到初始目录

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                filePath = openFileDialog.FileName;
            }

            string installPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            string extractPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"\Cache\";
            if (Directory.Exists(extractPath)) Directory.Delete(extractPath,true);
            Directory.CreateDirectory(extractPath);

            ZipFile.ExtractToDirectory(filePath, extractPath);

            var file=extractPath + @"MoveToExtension.dll";
            var version_new = FileVersionInfo.GetVersionInfo(file).FileVersion;
            var version_old = ReadVersion();
            if (String.Compare(version_new, version_old)<0) return;

            if (!UnInstall(out string msg))
            {
                MessageBox.Show(msg);
                return;
            }
            RestartExplorer();

            CopyDirectory(extractPath, installPath);

            if (!Install(out string msg1))
            {
                MessageBox.Show(msg1);
                return;
            }
            RestartExplorer();

            label_version.Text = "版本号:" + version_new;
            DeleteVersion();
            WriteVersion(version_new);
            MessageBox.Show("升级完成");
            btn_install.Enabled = false;
            btn_uninstall.Enabled = true;
        }

        public void CopyDirectory(string sourceDir, string targetDir)
        {
            // 创建目标目录
            Directory.CreateDirectory(targetDir);

            // 复制所有文件
            foreach (string file in Directory.EnumerateFiles(sourceDir))
            {
                if (Path.GetFileName(file).ToLower().Contains("install")) continue;
                string destFile = Path.Combine(targetDir, Path.GetFileName(file));
                File.Copy(file, destFile, true); // 覆盖已存在的文件
            }

            // 递归复制所有子目录
            foreach (string dir in Directory.EnumerateDirectories(sourceDir))
            {
                string destDir = Path.Combine(targetDir, Path.GetFileName(dir));
                CopyDirectory(dir, destDir);
            }
        }
    }
}
