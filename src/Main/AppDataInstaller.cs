using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
using System.Security.Principal;
using System.Collections.Generic;
using System.Xml;

namespace USC.GISResearchLab.Common.WebDeploymentHelper
{
    [RunInstaller(true)]
    public partial class WebDeploymentHelper : Installer
    {
        public WebDeploymentHelper()
        {
            InitializeComponent();
        }

        public static void Main()
        {
            WebDeploymentHelper.Install(@"?^InstallFolder^C:\inetpub\wwwroot\?WriteFolder^..\Data?ConfigFile^web.config?ShortestPathService.shortestpathservice^[URL]?DatabaseDirectory^[DBFOLDER]?EmailConnectionString^[EMAIL]?WebGISToolsConnectionString^[SQL1]?WebGISUserDatabaseTablesConnectionString^[SQL2]?ParallelRunner^[PR]");
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);
            Install(this.Context.Parameters["params"]);
        }

        private static void Install(string p)
        {
            string WriteFolder = string.Empty, ConfigFile = string.Empty, tempDir = string.Empty;
            string user = string.Empty;
            var keyValue = WebDeploymentHelper.ParseParameters(p);
            bool Changed = false;

            if (keyValue.ContainsKey("InstallFolder") && Directory.Exists(keyValue["InstallFolder"]))
            {
                tempDir = Environment.CurrentDirectory;
                Environment.CurrentDirectory = keyValue["InstallFolder"];
                if (keyValue.ContainsKey("WriteFolder"))
                {
                    WriteFolder = Path.GetFullPath(keyValue["WriteFolder"]);
                    Directory.CreateDirectory(WriteFolder);
                    // Try add access to the AppData folder for the Users user
                    // Add the access control entry to the directory.
                    // Other users worth knowing about
                    // user = Environment.MachineName + "\\ASPNET";
                    // user = Environment.MachineName + "\\IIS_IUSRS";
                    // user = new SecurityIdentifier(WellKnownSidType.NetworkServiceSid, null).Translate(typeof(NTAccount)).Value.ToString();
                    try
                    {
                        user = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null).Translate(typeof(NTAccount)).Value.ToString();
                        AddDirectorySecurity(WriteFolder, user, FileSystemRights.Modify, AccessControlType.Allow);
                        AddDirectorySecurity(WriteFolder, user, FileSystemRights.Write, AccessControlType.Allow);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error on adding security for user: " + user + ", folder: " + WriteFolder + Environment.NewLine + ex.Message);
                        throw ex;
                    }
                }

                if (keyValue.ContainsKey("ConfigFile"))
                {
                    ConfigFile = keyValue["ConfigFile"];
                    ConfigFile = Path.GetFullPath(ConfigFile);
                    if (File.Exists(ConfigFile))
                    {
                        // Web.config modification
                        try
                        {
                            XmlDocument doc = new XmlDocument();
                            doc.Load(ConfigFile);

                            XmlNodeList nodeList = null;
                            XmlNode root = doc.DocumentElement;

                            foreach (var c in keyValue)
                            {
                                if ((c.Key != "WriteFolder") && (c.Key != "ConfigFile") && (c.Key != "InstallFolder"))
                                {
                                    nodeList = root.SelectNodes("appSettings/add[@key='" + c.Key + "']");
                                    if (nodeList.Count > 0)
                                    {
                                        nodeList[0].Attributes["value"].InnerText = c.Value;
                                        Changed = true;
                                    }
                                    else
                                    {
                                        nodeList = root.SelectNodes("connectionStrings/add[@name='" + c.Key + "']");
                                        if (nodeList.Count > 0)
                                        {
                                            nodeList[0].Attributes["connectionString"].InnerText = c.Value;
                                            Changed = true;
                                        }
                                    }

                                }
                            } // end foreach. key/value updated.

                            if (Changed) doc.Save(ConfigFile);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error on updating config file:" + Environment.NewLine + ex.Message);
                            throw ex;
                        }
                    }
                }
                Environment.CurrentDirectory = tempDir;
            }
        }

        public override void Commit(System.Collections.IDictionary savedState)
        {
            base.Commit(savedState);
        }

        public override void Rollback(System.Collections.IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            base.Uninstall(savedState);
        }

        /// <summary>
        /// Adds an ACL entry on the specified directory for the specified account.
        /// </summary>
        /// <param name="dirName">Folder path to add the permissions to</param>
        /// <param name="Account"></param>
        /// <param name="Rights"></param>
        /// <param name="ControlType"></param>
        private static void AddDirectorySecurity(string folderName, string account, FileSystemRights rights, AccessControlType controlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(folderName);

            // Get a DirectorySecurity object that represents the current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();
            
            // Add the FileSystemAccessRule to the security settings. 
            dSecurity.AddAccessRule(new FileSystemAccessRule(account, rights, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, controlType));

            // Set the new access settings.
            dInfo.SetAccessControl(dSecurity);
        }

        private static Dictionary<string, string> ParseParameters(string input)
        {
            var ret = new Dictionary<string, string>();
            char c1 = char.MinValue, c2 = char.MinValue;
            string[] temp = null, list = null;
            try
            {
                c1 = input[0];
                c2 = input[1];
                input = input.Substring(2);
                list = input.Split(c1);

                foreach (var s in list)
                {
                    temp = s.Split(c2);
                    ret.Add(temp[0], temp[1]);
                }
            }
            catch { }
            return ret;
        }
    }
}