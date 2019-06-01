using System;
using System.Diagnostics;
using System.IO;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RP_FolderPinning.Model;
namespace RP_FolderPinning
{
    class Program
    {
        private static Config config;
        private static bool verbose = false;
        private static int delay = 0;
        private static string current_directory = Directory.GetCurrentDirectory();
        static void Main(string[] args)
        {
            Console.Title = "Lesson Pinning Created by Firman Jamal";
            LoadJson();
            verbose = config.verbose;
            if (verbose)
            {
                delay = config.verbose_sleep_time * 1000;
            }
            if (!Directory.Exists("Modules"))
            {
                string[] folders = Directory.GetDirectories(config.folder_dir);
                Directory.CreateDirectory("Modules");
                foreach (var folder in folders)
                {
                    string name = folder.Replace(config.folder_dir + "\\", "");
                    if (name.Length > 3)
                    {
                        name = name.Substring(0, 4);
                    }
                    CreateShortcut(name, folder);
                }
            }
            if (!String.IsNullOrEmpty(config.prev_dir))
            {
                Console.WriteLine("Removing previous lesson...");
                Process process = new Process();
                process.StartInfo.FileName = current_directory + "\\syspin.exe";
                process.StartInfo.Arguments = config.prev_dir + " 5387";
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                SaveJson("");
                System.Threading.Thread.Sleep(3000);
            }
            if (verbose == true)
            {
                Console.WriteLine("-------------------------------------");
            }
            

            foreach (var calendar in config.calendar)
            {
                DateTime dtClass = TimeStampToDateTime(calendar.StartTime).Date;
                DateTime now = DateTime.Now.Date;
                Console.WriteLine(calendar.Title);
                if (DateTime.Compare(now, dtClass) == 0)
                {
                    if (verbose == true)
                    {
                        Console.WriteLine("-------------------------------------");
                        Console.WriteLine("Today's Date " + DateTime.Now.ToString("dddd, dd MMMM yyyy"));
                        Console.WriteLine("-------------------------------------");
                        Console.WriteLine("");
                        Console.WriteLine("Module Code " + calendar.Title);
                        Console.WriteLine("Venue " + calendar.Venue);
                    }
                    Console.WriteLine("Pinning current lesson...");
                    string code = calendar.Title.Substring(0, 4);
                    string directory = current_directory + "\\Modules\\" + code + ".lnk";
                    Process process = new Process();
                    process.StartInfo.FileName = current_directory + "\\syspin.exe";
                    process.StartInfo.Arguments = directory + " 5386";
                    process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    process.Start();
                    SaveJson(directory);
                    break;
                }
            }
            if (verbose == true)
            {
                System.Threading.Thread.Sleep(3000);
            }
        }

        private static void CreateShortcut(string name, string path)
        {
            WshShell wsh = new WshShell();
            IWshShortcut shortcut = wsh.CreateShortcut(current_directory + "\\Modules\\" + name + ".lnk") as IWshRuntimeLibrary.IWshShortcut;
            shortcut.TargetPath = "C:\\Windows\\explorer.exe";
            shortcut.Arguments = path;
            shortcut.WindowStyle = 1;
            shortcut.Save();
        }

        private static void LoadJson()
        {
            if (!System.IO.File.Exists("config.json"))
            {
                string json = @"
                    {
                      'calendar': [
                      ],
                      'prev_dir': '',
                      'folder_dir': '',
                      'verbose': false,
                      'verbose_sleep_time': 3
                    }
                ";
                using (StreamWriter sw = new StreamWriter("config.json"))
                {
                    sw.WriteLine(json);
                    sw.Close();
                }
            }
            using (StreamReader reader = new StreamReader("config.json"))
            {
                string data = reader.ReadToEnd();
                config = JsonConvert.DeserializeObject<Config>(data);
                reader.Close();
            }
        }

        private static void SaveJson(string prev_dir)
        {
            JObject jsonD;
            using (StreamReader rd = new StreamReader("config.json"))
            {
                string data = rd.ReadToEnd();
                jsonD = JObject.Parse(data);
                jsonD["prev_dir"] = prev_dir;
                rd.Close();
            }
            using (StreamWriter sw = new StreamWriter("config.json"))
            {
                sw.WriteLine(jsonD);
                sw.Close();
            }
        }

        private static DateTime TimeStampToDateTime(double ts)
        {
            System.DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dt = dt.AddMilliseconds(ts).ToLocalTime();
            return dt;
        }
    }
}
