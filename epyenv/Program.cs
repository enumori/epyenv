
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

using System.Text.RegularExpressions;
using System.IO.Compression;
using HtmlAgilityPack;
using System.Diagnostics;
using System.Windows.Forms;
using CommandLine;
using Alphaleonis;

namespace epyenv
{
    public class Program
    {
        public class VersionInfo
        {
            public string Version { set; get; }
            public string Arch { set; get; }

            public string FileName { set; get; }

            public string VersionName
            {
                get
                {
                    return Version + "-" + Arch;
                }
            }
        }
        static string _AppDir = "";
        static string _InstallDir = "";
        static int _WaitCharId = 0;
        public enum ErrorCode
        {
            NO_ERROR,
            INVALID_ARGS,
            WITH_NOT_PARSED
        }


        [Verb("--version", HelpText = "epyenvのバージョンを出力します。")]
        public class ShowVersion
        {
        }

        [Verb("install", HelpText = "指定されたバージョンのPython Windows embeddable packageをインストールします。")]
        public class Install
        {
            [CommandLine.Value(0, MetaName = "Version", Required = true, HelpText = "インストールするバージョン")]
            public string Version { get; set; }
            [CommandLine.Option('o', "out", Required = false, HelpText = "インストール先ディレクトリパス")]
            public string Dir { get; set; }
        }
        [Verb("list", HelpText = "インストール可能なバージョンを出力します。")]
        public class VersionList
        {
        }
        static void Main(string[] args)
        {
            ErrorCode error_code = ErrorCode.INVALID_ARGS;
            _AppDir = System.AppDomain.CurrentDomain.BaseDirectory;
            _InstallDir = System.IO.Directory.GetCurrentDirectory();

            var parser = new Parser(config => { config.IgnoreUnknownArguments = false; config.AutoVersion = false; config.HelpWriter = Console.Out; });
            var result = parser.ParseArguments<ShowVersion, Install, VersionList, Version>(args)
                .WithParsed<ShowVersion>(opts => {
                    error_code = ProcShowVersion();
                })
                .WithParsed<Install>(opts => {
                    error_code = ProcInstall(opts);
                })
                .WithParsed<VersionList>(opts => {
                    error_code = ProcVersionList();
                })
                .WithNotParsed(errs => {
                    error_code = ErrorCode.WITH_NOT_PARSED;
                });

            if (error_code == ErrorCode.INVALID_ARGS)
            {
                parser.ParseArguments<ShowVersion, Install, VersionList, Version>(new string[] { "" });
            }

#if DEBUG
            Console.WriteLine("続行するには何かキーを押してください．．．");
            Console.ReadKey();
#endif
            Environment.Exit((error_code != ErrorCode.NO_ERROR) ? 1 : 0);
        }

        static ErrorCode ProcShowVersion()
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.Console.WriteLine(ver.FileVersion);
            return ret_code;
        }
        static ErrorCode ProcInstall(Program.Install opt)
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            System.Diagnostics.Debug.WriteLine("ProcInstall");
            System.Diagnostics.Debug.WriteLine("opt.Version:" + opt.Version);
            System.Diagnostics.Debug.WriteLine("opt.Dir:" + opt.Dir);
            if (opt.Dir != null)
            {
                _InstallDir = GetFullPath(System.IO.Directory.GetCurrentDirectory() + "\\", opt.Dir);
            }
            if (opt.Version != null)
            {
                ret_code = InstallPythonEmbed(opt.Version);
            }
            return ret_code;
        }
        static string WaitChar()
        {

            if (_WaitCharId == 0)
            {
                _WaitCharId = 1;
                return "\b|";
            }
            else
            {
                _WaitCharId = 0;
                return "\b-";
            }

        }

        static ErrorCode ProcVersionList()
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            string url = Properties.Settings.Default.PythonURL;
            string source = "";
            List<VersionInfo> version_info_list = new List<VersionInfo>();
            List<string> version_list = new List<string>();
            System.Net.HttpWebRequest webreq = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            Console.Write(url + "を検索中です。 ");
            using (System.Net.HttpWebResponse webres = (System.Net.HttpWebResponse)webreq.GetResponse())
            {
                using (System.IO.Stream st = webres.GetResponseStream())
                {
                    //文字コードを指定して、StreamReaderを作成
                    using (System.IO.StreamReader sr = new System.IO.StreamReader(st, System.Text.Encoding.UTF8))
                    {
                        source = sr.ReadToEnd();
                    }
                }
            }

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(source);
            foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
            {
                Console.Write(WaitChar());
                if (System.Text.RegularExpressions.Regex.IsMatch(link.InnerHtml, @"^[0-9]+(\.[0-9]+)*"))
                {
                    Regex re = new Regex(@"^[0-9]+(\.[0-9]+)*/", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    Match m = re.Match(link.InnerHtml);
                    while (m.Success)
                    {
                        Console.Write(WaitChar());
                        if (m.Value == link.InnerHtml)
                        {
                            version_list.Add(m.Value.TrimEnd('/'));
                        }
                        m = m.NextMatch();
                    }
                }                
            }
            foreach (var ver in version_list)
            {
                webreq = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url + '/' + ver + '/');
                using (System.Net.HttpWebResponse webres = (System.Net.HttpWebResponse)webreq.GetResponse())
                {
                    using (System.IO.Stream st = webres.GetResponseStream())
                    {
                        //文字コードを指定して、StreamReaderを作成
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(st, System.Text.Encoding.UTF8))
                        {
                            source = sr.ReadToEnd();
                        }
                    }
                }

                doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(source);
                foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
                {
                    Console.Write(WaitChar());
                    string file_name = "";
                    file_name = string.Format(Properties.Settings.Default.PythonEmbedName, ver, Properties.Settings.Default.ArchAmd64);
                    if (link.InnerHtml.Equals(file_name))
                    {
                        version_info_list.Add(new VersionInfo() { Version = ver, Arch = Properties.Settings.Default.ArchAmd64, FileName = file_name });
                    }
                    file_name = string.Format(Properties.Settings.Default.PythonEmbedName, ver, Properties.Settings.Default.ArchWin32);
                    if (link.InnerHtml.Equals(file_name))
                    {
                        version_info_list.Add(new VersionInfo() { Version = ver, Arch = Properties.Settings.Default.ArchWin32, FileName = file_name });
                    }
                }
            }
            Console.Write("\b ");
            Console.Write(Environment.NewLine);
            version_info_list.Sort((a, b) => string.Compare(Version2String(a.Version), Version2String(b.Version)));

            foreach (var ver in version_info_list)
            {
                Console.WriteLine(ver.VersionName);
            }
            return ret_code;
        }
        public static string GetFullPath(string base_path, string path)
        {
            string full_path = null;

            if (!System.IO.Path.IsPathRooted(path))
            {
                Uri u1 = new Uri(base_path);
                Uri u2 = new Uri(u1, path);
                full_path = u2.LocalPath;
            }
            else
            {
                full_path = path;
            }
            return full_path;
        }

        static void RemoveReadonlyAttribute(Alphaleonis.Win32.Filesystem.DirectoryInfo dirInfo)
        {
            if ((dirInfo.Attributes & System.IO.FileAttributes.ReadOnly) ==　System.IO.FileAttributes.ReadOnly)
            {
                dirInfo.Attributes = System.IO.FileAttributes.Normal;
            }
            foreach (var fi in dirInfo.GetFiles())
            {
                if ((fi.Attributes & System.IO.FileAttributes.ReadOnly) == System.IO.FileAttributes.ReadOnly)
                {
                    fi.Attributes = System.IO.FileAttributes.Normal;
                }
            }
            foreach (var di in dirInfo.GetDirectories())
            {
                RemoveReadonlyAttribute(di);
            }
            return;
        }

        static ErrorCode InstallPythonEmbed(string version)
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            string[] ver_str = version.Split('-');
            if (ver_str.Length != 2)
            {
                System.Console.Error.WriteLine("不正なバージョン文字列です。");
                return ErrorCode.INVALID_ARGS;
            }
            VersionInfo ver = new VersionInfo() { Version = ver_str[0], Arch = ver_str[1], FileName = "" };
            if (!System.Text.RegularExpressions.Regex.IsMatch(ver.Version, @"^[0-9]+(\.[0-9]+)*"))
            {
                System.Console.Error.WriteLine("不正なバージョン文字列です。");
                return ErrorCode.INVALID_ARGS;
            }
            ver.FileName = string.Format(Properties.Settings.Default.PythonEmbedName, ver.Version, ver.Arch);

            if ((ret_code = DownloadPythonEmbed(ver)) != ErrorCode.NO_ERROR)
            {
                return ret_code;
            }
            if ((ret_code = DownloadPip()) != ErrorCode.NO_ERROR)
            {
                return ret_code;
            }
            if ((ret_code = ExtractTkinter(ver)) != ErrorCode.NO_ERROR)
            {
                return ret_code;
            }
            return ret_code;
        }

        static string Version2String(string version)
        {
            Regex re = new Regex(@"[0-9]+", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = re.Match(version);
            string ver_string = "";
            while (m.Success)
            {
                ver_string += string.Format("{0:D4}", Convert.ToInt32(m.Value));
                m = m.NextMatch();
            }
            return ver_string;
        }

        static ErrorCode DownloadPythonEmbed(VersionInfo ver)
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;

            string python = System.IO.Path.Combine(_InstallDir, "python.exe");

            if (!System.IO.Directory.Exists(_InstallDir))
            {
                System.Diagnostics.Debug.WriteLine("ディレクトリ作成: " + _InstallDir);
                System.IO.Directory.CreateDirectory(_InstallDir);
            }
            if (!System.IO.File.Exists(python))
            {


                string url = string.Format("{0}/{1}/{2}", Properties.Settings.Default.PythonURL, ver.Version, ver.FileName);

                System.Net.WebClient wc = new System.Net.WebClient();


                string download = System.IO.Path.Combine(_InstallDir, ver.FileName);
                try
                {
                    wc.DownloadFile(url, download);
                    System.Console.WriteLine("Python Windows embeddable package(" + ver.FileName + ")をインストール中");
                    System.Diagnostics.Debug.WriteLine("ダウンロード: " + url + " -> " + ver.FileName);
                }
                catch
                {
                    System.Console.Error.WriteLine("指定されたバージョンが見つかりませんでした。");
                    return ErrorCode.INVALID_ARGS;
                }
                finally
                {
                    wc.Dispose();
                }
                ExtractToDirectoryExtensions(download, _InstallDir, true);
                System.IO.File.Delete(download);
            }
            return ret_code;
        }

        static ErrorCode DownloadPip()
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            string url = Properties.Settings.Default.GetPipURL;
            var download = System.IO.Path.Combine(_InstallDir, Properties.Settings.Default.GetPipName);
            var python = System.IO.Path.Combine(_InstallDir, "python.exe");

            string scripts_dir = System.IO.Path.Combine(_InstallDir, "Scripts");
            string pip = System.IO.Path.Combine(scripts_dir, "pip.exe");

            if (!System.IO.Directory.Exists(_InstallDir))
            {
                System.Diagnostics.Debug.WriteLine("ディレクトリ作成: " + _InstallDir);
                System.IO.Directory.CreateDirectory(_InstallDir);
            }
            if (!System.IO.File.Exists(pip))
            {
                System.Console.WriteLine("pipをインストールします。");
                System.Console.WriteLine("get-pipのダウンロード中");
                System.Net.WebClient wc = new System.Net.WebClient();
                wc.DownloadFile(url, download);
                wc.Dispose();

                System.Console.WriteLine("pipのインストール中");
                Process p;

                p = new Process();
                p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.WorkingDirectory = _InstallDir;
                p.StartInfo.Arguments = string.Format("/c \"\"{0}\" \"\"{1}\"", python, download);
                p.Start();
                p.WaitForExit();
                p.Close();
            }
            System.IO.File.Delete(download);
            return ret_code;
        }

        static ErrorCode ExtractTkinter(VersionInfo ver)
        {
            ErrorCode ret_code = ErrorCode.NO_ERROR;
            string tkinter_file = "";
            if (string.Compare(Version2String(ver.Version), Version2String("3.6.0")) < 0)
            {
                if (ver.Arch == Properties.Settings.Default.ArchAmd64)
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_5_Amd64);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_5_Amd64));
                }
                else
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_5_Win32);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_5_Win32));
                }
                ExtractToDirectoryExtensions(tkinter_file, _InstallDir, true);
            }
            else if (string.Compare(Version2String(ver.Version), Version2String("3.7.0")) < 0)
            {
                if (ver.Arch == Properties.Settings.Default.ArchAmd64)
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_6_Amd64);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_6_Amd64));
                }
                else
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_6_Win32);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_6_Win32));
                }
                ExtractToDirectoryExtensions(tkinter_file, _InstallDir, true);
            }
            else if (string.Compare(Version2String(ver.Version), Version2String("3.8.0")) < 0)
            {
                if (ver.Arch == Properties.Settings.Default.ArchAmd64)
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_7_Amd64);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_7_Amd64));
                }
                else
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_7_Win32);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_7_Win32));
                }
                ExtractToDirectoryExtensions(tkinter_file, _InstallDir, true);
            }
            else if (string.Compare(Version2String(ver.Version), Version2String("3.9.0")) < 0)
            {
                if (ver.Arch == Properties.Settings.Default.ArchAmd64)
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_8_Amd64);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_8_Amd64));
                }
                else
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_8_Win32);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_8_Win32));
                }
                ExtractToDirectoryExtensions(tkinter_file, _InstallDir, true);
            }
            else
            {
                if (ver.Arch == Properties.Settings.Default.ArchAmd64)
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_9_Amd64);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_9_Amd64));
                }
                else
                {
                    tkinter_file = System.IO.Path.Combine(_AppDir, Properties.Settings.Default.tkinter3_9_Win32);
                    System.Console.WriteLine(string.Format("tkinter({0})をインストール中", Properties.Settings.Default.tkinter3_9_Win32));
                }
                ExtractToDirectoryExtensions(tkinter_file, _InstallDir, true);
            }
            return ret_code;
        }

        static void ExtractToDirectoryExtensions(string sourceArchiveFileName, string destinationDirectoryName, bool overwrite)
        {
            using (ZipArchive archive = ZipFile.OpenRead(sourceArchiveFileName))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    var fullPath = System.IO.Path.Combine(destinationDirectoryName, entry.FullName);
                    if (string.IsNullOrEmpty(entry.Name))
                    {
                        if (!System.IO.Directory.Exists(fullPath))
                        {
                            System.IO.Directory.CreateDirectory(fullPath);
                        }
                    }
                    else
                    {
                        if (overwrite)
                        {
                            entry.ExtractToFile(fullPath, true);
                        }
                        else
                        {
                            if (!System.IO.File.Exists(fullPath))
                            {
                                entry.ExtractToFile(fullPath, true);
                            }
                        }
                    }
                }
            }
        }

        static bool ConsoleQuestion(string question)
        {
            string ans = "";
            bool ret = false;
            ConsoleKeyInfo response;
            while (true)
            {
                Console.Write(question);
                while (true)
                {
                    response = Console.ReadKey(true);
                    if (response.Key == ConsoleKey.Enter)
                    {
                        break;
                    }
                    else
                    {
                        Console.Write(response.KeyChar);
                        ans += response.KeyChar;
                    }
                }
                Console.WriteLine();
                if (ans.Equals("Y", StringComparison.InvariantCultureIgnoreCase)
                    || ans.Equals("N", StringComparison.InvariantCultureIgnoreCase))
                {
                    break;
                }
                else
                {
                    ans = "";
                }
            }
            if (ans.Equals("Y", StringComparison.InvariantCultureIgnoreCase))
            {
                ret = true;
            }
            return ret;
        }
    }
}
