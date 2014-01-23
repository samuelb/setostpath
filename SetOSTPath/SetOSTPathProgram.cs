namespace SetOSTPath
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using Microsoft.Win32;

    /// <summary>
    /// This program sets the .ost file locate of a MS Outlook profile.  
    /// </summary>
    public class SetOSTPathProgram
    {
        /// <summary>
        /// Entry point of the program
        /// </summary>
        /// <param name="args">Command line arguments</param>
        public static void Main(string[] args)
        {
            string profile = null;
            string newostpath = null;
            
            // check for possible help arguments
            foreach (var arg in args) {
                switch (arg.ToLower()) {
                    case "-h":
                    case "-?":
                    case "-help":
                    case "--help":
                    case "/h":
                    case "/?":
                    case "/help":
                        printUsage();
                        Environment.Exit(0);
                        break;
                }
            }
            
            if (args.Length == 0) {
                printUsage();
                Environment.Exit(0);
            } else {
                profile = args[0];
            }

            if (args.Length >= 2) {
                 newostpath = args[1]; 
                 //newostpath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) , @"Microsoft\Outlook\oulook.ost");
            }
            
            try {
                var start = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\" + profile);
                if (start == null) {
                    Console.Error.WriteLine("No profile named '" + profile + "' found");
                    Environment.Exit(1);
                }
                
                var data = SearchReg(start, "001f6610");
                var regentry = data.Key;
                var ostpath = Encoding.Unicode.GetString(data.Value as byte[]);
                Console.WriteLine("Current OST Path: " + ostpath.Trim(new char[]{'\0', '\n', '\r'}));
                
                if (newostpath != null && newostpath.Length > 0) {
                    Console.WriteLine("New OST Path: " + newostpath);
                    var regentrydir = Path.GetDirectoryName(regentry);
                    regentrydir = regentrydir.Remove(0,    regentrydir.IndexOf('\\') + 1); // because .Split('\\', 1)[1] not working
                    var newostpathdata = Encoding.Unicode.GetBytes(newostpath + "\0");
                    Registry.CurrentUser.OpenSubKey(regentrydir, true).SetValue("001f6610", newostpathdata);
                }
            } catch (KeyNotFoundException) {
                Console.Error.WriteLine("OST Path not found on registry");
                Environment.Exit(2);
            } catch (Exception e) {
                Console.Error.WriteLine("Error: " + e.Message);
                Environment.Exit(3);
            }
        }
        
        /// <summary>
        /// Search the registry for the given key
        /// </summary>
        /// <param name="key">Registry key name to search</param>
        /// <param name="search">Scope to search within</param>
        /// <returns>Pair with found key and the value of it</returns>
        public static KeyValuePair<string, object> SearchReg(RegistryKey key, string search) {
            Queue<RegistryKey> toSearch = new Queue<RegistryKey>();
            if (key != null) toSearch.Enqueue(key);
            
            while (toSearch.Count > 0) {
                var currentkey = toSearch.Dequeue();
                foreach (var val in currentkey.GetValueNames()) {
                    if (val == search) {
                        return new KeyValuePair<string, object>(currentkey.Name + @"\" + val, currentkey.GetValue(val));
                    }
                }
                foreach (var subkey in currentkey.GetSubKeyNames()) {
                    toSearch.Enqueue(currentkey.OpenSubKey(subkey));
                }
            }
            throw new KeyNotFoundException();
        }
        
        /// <summary>
        /// Prints a brief usage info to the terminal
        /// </summary>
        public static void printUsage() {
            Console.WriteLine("Usage: profile ost-path");
        }

    } 
}