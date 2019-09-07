using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Main
{
    class Mapping
    {
        private string _sid;

        public string Sid
        {
            get { return _sid; }
            set
            {
                _sid = value;
                try
                {
                    Username = new System.Security.Principal.SecurityIdentifier(_sid).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Invalid Username...{ex}");
                    Username = _sid;
                }
            }
        }

        public string Username;
        public override string ToString()
        {
            return $"{Username} : {Path}";
        }
        public string Path;


        public string Version { get; set; }
    }
    class Program
    {
        const string OST_PATH = @"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook";

        static readonly Dictionary<string, string> PSTPaths = new Dictionary<string, string>
        {
            ["2007"] = @"software\Microsoft\Office\12.0\Outlook",
            ["2010"] = @"software\Microsoft\Office\14.0\Outlook",
            ["2013"] = @"software\Microsoft\Office\15.0\Outlook",
            ["2016"] = @"software\Microsoft\Office\16.0\Outlook",
        };


        const string OST_KEY = "001f6610";
        static List<Mapping> result = new List<Mapping>();
        static Dictionary<string, string> ProfilePathSidDict = new Dictionary<string, string>();
        static void Main(string[] args)
        {

            bool isXP = System.Environment.OSVersion.Version.Major == 5;
            string user_path = isXP ? "C:\\Documents and Settings" : "C:\\Users";



            //construct profile key..
            var profileSubKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\ProfileList");
            foreach (var key in GetKeys(profileSubKey))
            {
                var profileKey = profileSubKey.OpenSubKey(key);
                var profilePath = (string)profileKey.GetValue("ProfileImagePath");
                string resolvedVal = key;
                try
                {
                    resolvedVal = new System.Security.Principal.SecurityIdentifier(key).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
                }
                catch (Exception ex)
                {

                }
                ProfilePathSidDict[profilePath] = resolvedVal;
            }



            string filePath = "./";
            if (args != null && args.Length > 0)
            {
                filePath = args[0];
            }

            OnlineUsersGetOstAndPst();

            if (Directory.Exists(user_path))
            {
                string[] sub_dirs = Directory.GetDirectories(user_path);
               // int i = 0;
                foreach (string dir in sub_dirs)
                {
                  //  RegistryInterop.Unload("test" + i);

                    var domainName = dir;
                    if (ProfilePathSidDict.ContainsKey(dir))
                    {
                        domainName = ProfilePathSidDict[dir];
                    }
                    Console.WriteLine("Directory is:" + dir);

                    string wimHivePath = Path.Combine(dir, "ntuser.dat");
                    string copy = Path.Combine(dir, "ntuserCOPY.dat");

                    try
                    {

                        File.Copy(wimHivePath, copy, true);

                        string loadedHiveKey = RegistryInterop.Load(copy);
                        if (string.IsNullOrEmpty(loadedHiveKey))
                        {
                            RegistryInterop.Unload();
                            continue;
                        }

                        RegistryKey rk = Microsoft.Win32.Registry.Users.OpenSubKey(loadedHiveKey);

                        if (rk != null)
                        {
                            ReadFromRootKey(rk, domainName);
                            rk.Close();

                        }

                        File.Delete(copy);
                    }
                    catch (Exception ex)
                    {
                        RegistryInterop.Unload();
                        //  Console.WriteLine(ex);
                        continue;
                    }

                    RegistryInterop.Unload();

                }
            }

            var machineName = System.Environment.MachineName;



            Console.WriteLine("---------------------------------------------");

            var csv = new StringBuilder();
            //contains all ost and pst..
            foreach (var ss in result)
            {
                Console.WriteLine(ss);
                var s1 = $"{machineName},{ss.Username},{ss.Path},{ss.Version}";
                csv.AppendLine(s1);
            }
            System.IO.File.WriteAllText($"{filePath}\\{machineName}.csv", csv.ToString());

            Console.WriteLine("Done.");
        }

        public static void OnlineUsersGetOstAndPst()
        {
            RegistryKey rk = RegistryKey.OpenRemoteBaseKey(RegistryHive.Users, "");
            foreach (var s in rk.GetSubKeyNames())
            {
                var userRootKey = rk.OpenSubKey(s);
                if (userRootKey != null)
                {
                    ReadFromRootKey(userRootKey, s);
                    userRootKey.Close();
                }
            }
            rk.Close();

        }

        public static void ReadFromRootKey(RegistryKey rootKey, string sid)
        {

            if (rootKey != null)
            {
                var ostRootKey = rootKey.OpenSubKey(OST_PATH);
                if (ostRootKey != null)
                {
                    foreach (var path in GetKeys(ostRootKey))
                    {
                        var tmpList = ostRootKey.OpenSubKey(path);
                        var ostPathArr = tmpList.GetValue(OST_KEY);
                        if (ostPathArr != null)//OST PATH found, convert it 
                        {
                            //Unicode..
                            var realPath = System.Text.Encoding.Unicode.GetString((byte[])ostPathArr);
                            //Console.WriteLine(realPath);
                            result.Add(new Mapping { Sid = sid, Path = realPath });
                            continue;
                        }
                        tmpList.Close();
                    }
                    ostRootKey.Close();
                }

                //Get PST
                foreach (var kvp in PSTPaths)
                {
                    var path = kvp.Value;
                    var tmpRootKey = rootKey.OpenSubKey(path);
                    if (tmpRootKey != null)
                    {
                        var _result = OutputRegKey(tmpRootKey).Where(x => x.Contains("pst")).ToList();
                        if (_result != null)
                        {
                            foreach (var a in _result.Distinct())
                            {
                                result.Add(new Mapping { Sid = sid, Path = a, Version = kvp.Key });
                            }
                        }
                        tmpRootKey.Close();
                    }
                };
            }



        }

        private static List<string> processValueNames(RegistryKey Key)
        { //function to process the valueNames for a given key
            string[] valuenames = Key.GetValueNames();
            if (valuenames == null || valuenames.Length <= 0) //has no values
                return new List<string>();
            return valuenames.ToList();
        }

        public static List<string> OutputRegKey(RegistryKey Key)
        {
            try
            {
                List<string> resultList = new List<string>();

                string[] subkeynames = Key.GetSubKeyNames(); //means deeper folder
                if (subkeynames == null || subkeynames.Length <= 0)
                { //has no more subkey, process
                    resultList.AddRange(processValueNames(Key));
                    return resultList;
                }
                foreach (string keyname in subkeynames)
                { //has subkeys, go deeper
                    using (RegistryKey key2 = Key.OpenSubKey(keyname))
                        resultList.AddRange(OutputRegKey(key2));
                }
                if (processValueNames(Key) != null)
                {
                    resultList.AddRange(processValueNames(Key));
                }
                Key.Close();
                return resultList;
            }
            catch (Exception e)
            {
                // Console.WriteLine("Exception...", e);
            }
            Key.Close();
            return new List<string>();
        }

        static List<string> GetKeys(RegistryKey rkey, int limit = 100)
        {

            // Retrieve all the subkeys for the specified key.
            String[] names = rkey.GetSubKeyNames();

            int icount = 0;

            var resultList = new List<String>();
            // Print the contents of the array to the console.
            foreach (String s in names)
            {
                resultList.Add(s);
                // The following code puts a limit on the number
                // of keys displayed.  Comment it out to print the
                // complete list.
                icount++;
                //if (icount >= limit)
                //    break;
            }
            return resultList;
        }
    }
}
