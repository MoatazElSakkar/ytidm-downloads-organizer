using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Shell;
using Newtonsoft.Json;

namespace YtIdmDownloadsFolder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class YtIdmDownloadsFolderProcess
    {
        private Timer regPollTimer;
        private Timer organizerTimer;
        private RegistryKey IdmKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\DownloadManager",false);
        private const string ytVideoEmbedUrl = @"https://noembed.com/embed?url={0}";
        private long pollIndexOffset=1;
        FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
        private Dictionary<string, YtVideoInfo> fileInfoDictionry = new Dictionary<string, YtVideoInfo>();
        private List<FileSystemEventArgs> filesDownloadEvents = new List<FileSystemEventArgs>();
        private string currentPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        public YtIdmDownloadsFolderProcess()
        {
            InitializeComponent();

        }

        private void OrganizerTimerOnElapsed(object sender, ElapsedEventArgs eventArgs)
        {
            organizerTimer.Enabled = false;
            try
            {
                for (int i = 0; i < filesDownloadEvents.Count; i++)
                {
                    FileSystemEventArgs e = filesDownloadEvents[i];

                    string key = fileInfoDictionry.Keys.FirstOrDefault(x =>
                        e.Name.RemoveSpecialCharacters().Contains(x.RemoveSpecialCharacters()));
                    if (key != null)
                    {
                        if (!Directory.Exists(fileSystemWatcher.Path + "\\" + fileInfoDictionry[key].author_name))
                            Directory.CreateDirectory(
                                fileSystemWatcher.Path + "\\" + fileInfoDictionry[key].author_name);

                        File.Move(e.FullPath.Replace(".tmp", ""),
                            fileSystemWatcher.Path + "\\" + fileInfoDictionry[key].author_name.ValidateFileName() +
                            "\\" +
                            fileInfoDictionry[key].title.ValidateFileName() + ".mp4");
                        ShellFile shellFile = ShellFile.FromFilePath(fileSystemWatcher.Path + "\\" +
                                                                     fileInfoDictionry[key].author_name
                                                                         .ValidateFileName()
                                                                     + "\\" + fileInfoDictionry[key].title
                                                                         .ValidateFileName() + ".mp4");
                        shellFile.Properties.System.Comment.Value = fileInfoDictionry[key].url;
                        shellFile.Properties.System.Author.Value = new[] { fileInfoDictionry[key].author_name };
                        filesDownloadEvents.Remove(filesDownloadEvents[i]);
                        fileInfoDictionry.Remove(key);
                    }
                }
            }
            catch(Exception ex)
            {
                LogException(ex);
            }

            organizerTimer.Enabled = true;
        }
        
        private void FileSystemWatcherOnChanged(object sender, FileSystemEventArgs e)
        {
            FileSystemWatcherOnCreated(sender, e);
        }

        private async void RegPollTimerOnElapsed(object sender, ElapsedEventArgs e)
        {
            
            if (!GetIdmMaxId(out var pollMaxId)) 
                return;

            regPollTimer.Enabled = false;
            try{
                for (; pollIndexOffset < pollMaxId; pollIndexOffset++)
                {
                    RegistryKey itemKey = Registry.CurrentUser.OpenSubKey(
                        string.Format(@"SOFTWARE\DownloadManager\{0}", pollIndexOffset), false);
                    if (itemKey!=null)
                        continue;
                    
                    YtVideoInfo ytVideoInfo =
                        await AnalyzeYtVideoInfo(GetYtLink(itemKey));
                    
                    //                 if (File.Exists(ytVideoInfo.title + ".mp4"))
                    if (ytVideoInfo==null)
                        continue;
                    
                    if (fileInfoDictionry.ContainsKey(ytVideoInfo.title))
                        fileInfoDictionry[ytVideoInfo.title] = ytVideoInfo;
                    
                    else
                        fileInfoDictionry.Add(ytVideoInfo.title, ytVideoInfo);
                }
            }
            catch(Exception ex)
            {
                LogException(ex);
            }

            regPollTimer.Enabled = true;
        }

        private void FileSystemWatcherOnCreated(object sender, FileSystemEventArgs e)
        {
            filesDownloadEvents.Add(e);
        }

        private static bool GetIdmMaxId(out long pollMaxId)
        {
            pollMaxId=long.MaxValue;

            try
            {
                long.TryParse(
                    Registry.CurrentUser.OpenSubKey(@"SOFTWARE\DownloadManager\maxID", false)?
                        .GetValue("maxID").ToString(),
                    out pollMaxId);
                return true;
            }
            catch
            {
                Application.Current.Shutdown();
                return false;
            }

        }

        private async Task<YtVideoInfo> AnalyzeYtVideoInfo(string ytLink)
        {
            if (ytLink == null)
                return null;
            
            HttpWebRequest videoInfoWebRequest = (HttpWebRequest)WebRequest.Create(string.Format(ytVideoEmbedUrl,ytLink));
            videoInfoWebRequest.Method = "GET";
            WebResponse userinfoResponse = await videoInfoWebRequest.GetResponseAsync();
            using (StreamReader videoInfoResponseStream = new StreamReader(userinfoResponse.GetResponseStream()))
            {
                // reads response body
                string videoInfoResponse = await videoInfoResponseStream.ReadToEndAsync();
                return (YtVideoInfo)JsonConvert.DeserializeObject(videoInfoResponse, typeof(YtVideoInfo));
            }
        }

        private string GetYtLink(RegistryKey openSubKey)
        {
            string dlOwnerPage = openSubKey.GetValue("owWPage")?.ToString();
            if (dlOwnerPage != null && dlOwnerPage.StartsWith("https://www.youtube.com"))
                return dlOwnerPage;

            return null;
        }

        private async void YtIdmDownloadsFolderProcess_OnLoaded(object sender, RoutedEventArgs e)
        {
            regPollTimer = new Timer()
            {
                Interval = 30000,
                Enabled = true,
                AutoReset = true
            };            
            organizerTimer = new Timer()
            {
                Interval = 90000,
                Enabled = true,
                AutoReset = true
            };
            GetIdmMaxId(out pollIndexOffset);
            regPollTimer.Elapsed += RegPollTimerOnElapsed;
            organizerTimer.Elapsed+=OrganizerTimerOnElapsed;
            // if (File.Exists("presistent-poll-index-offset.dat"))
            //     long.TryParse(File.ReadAllText("presistent-poll-index-offset.dat"), out pollIndexOffset);
            fileSystemWatcher = new FileSystemWatcher();
            
            //The application should be in a folder in the youtube video folder
            fileSystemWatcher.Path = new DirectoryInfo(currentPath).Parent.Parent.FullName;
            
            fileSystemWatcher.Filter = "*.*";
            fileSystemWatcher.NotifyFilter = NotifyFilters.LastWrite;
            fileSystemWatcher.EnableRaisingEvents = true;
            fileSystemWatcher.Created += FileSystemWatcherOnCreated;
            fileSystemWatcher.Changed+= FileSystemWatcherOnChanged;
            
            //Performing some p/invoke magic to hide from alt/win tab
            WindowInteropHelper wndHelper = new WindowInteropHelper(this);
            int exStyle = (int)NativeMethods.GetWindowLong(wndHelper.Handle, (int)NativeMethods.GetWindowLongFields.GWL_EXSTYLE);
            exStyle |= (int)NativeMethods.ExtendedWindowStyles.WS_EX_TOOLWINDOW;
            NativeMethods.SetWindowLong(wndHelper.Handle, (int)NativeMethods.GetWindowLongFields.GWL_EXSTYLE, (IntPtr)exStyle);
        }
        
        private void LogException(Exception ex)
        {
            try
            {
                File.AppendAllLines(new DirectoryInfo(currentPath).Parent?.Parent + "\\failure.log",
                    new[]
                    {
                        DateTime.Now + "========================================================",
                        ex.Message,
                        ex.StackTrace,
                        "=====================================================================\n"
                    });
            }
            catch (Exception exi)
            {
                MessageBox.Show(exi.Message);
            }
        }
    }

}