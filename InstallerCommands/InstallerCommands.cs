using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace InstallerCommands
{
    [RunInstaller(true)]
    public partial class InstallerCommands : System.Configuration.Install.Installer
    {
        AddinType addinType = AddinType.Command;
        string commandProject = "RevitSimpleCommand";
        string commandName = "CmdHelloWorld";
        string companyName = "Lambda Ingenieria e Innovacion";
        string companyURL = "https://lambda.com.pe/";
        public InstallerCommands()
        {

            InitializeComponent();
        }
        /*
            ####### METODO DE INSTALACION #######
        */
        public override void Install(IDictionary stateSaver)
        {
            Microsoft.Win32.RegistryKey rkbase = null;
            rkbase = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
            rkbase.CreateSubKey($"SOFTWARE\\Wow6432Node\\{companyName}\\Revit API NuGet Example 2019 Packages", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree).SetValue("OokiiVersion", typeof(Ookii.Dialogs.Wpf.CredentialDialog).Assembly.FullName);
            rkbase.CreateSubKey($"SOFTWARE\\Wow6432Node\\{companyName}\\Revit API NuGet Example 2019 Packages", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree).SetValue("XceedVersion", typeof(Xceed.Wpf.Toolkit.PropertyGrid.PropertyGrid).Assembly.FullName);

            string sDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\Autodesk\\Revit\\Addins";
            bool exists = Directory.Exists(sDir);

            if (!exists) Directory.CreateDirectory(sDir);

            XElement XElementAddIn;
            if (addinType == AddinType.Command)
            {
                XElementAddIn = new XElement("AddIn", new XAttribute("Type", "Command"));
                XElementAddIn.Add(new XElement("Text", commandProject));
            }
            else
            {
                XElementAddIn = new XElement("AddIn", new XAttribute("Type", "Application"));
                XElementAddIn.Add(new XElement("Name", commandProject));
            }

            XElementAddIn.Add(new XElement("Assembly", this.Context.Parameters["targetdir"].Trim()  + commandProject + ".dll"));
            XElementAddIn.Add(new XElement("AddInId", Guid.NewGuid().ToString()));
            XElementAddIn.Add(new XElement("FullClassName", $"{commandProject}.{commandName}"));
            XElementAddIn.Add(new XElement("VendorId", "ADSK"));
            XElementAddIn.Add(new XElement("VendorDescription", $"{companyName}, {companyURL}"));

            XElement XElementRevitAddIns = new XElement("RevitAddIns");
            XElementRevitAddIns.Add(XElementAddIn);

            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    new XDocument(XElementRevitAddIns).Save(d + "\\" + commandProject + ".addin");
                }
            }
            catch (Exception excpt)
            {
                MessageBox.Show(excpt.Message);
            }
        }

    }
    enum AddinType
    {
        Command = 0,
        Application = 1
    }
}
