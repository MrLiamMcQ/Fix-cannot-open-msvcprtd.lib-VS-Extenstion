using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using System.IO;
using EnvDTE;
using EnvDTE80;
using System.Xml.Linq;
using System.Runtime.InteropServices;

namespace fixMsvcprtdBug
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class main_FixMsvcprtd
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("f39108ef-6124-4da9-af69-ba4d28b8f3a2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="main_FixMsvcprtd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private main_FixMsvcprtd(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static main_FixMsvcprtd Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in main_FixMsvcprtd's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new main_FixMsvcprtd(package, commandService);
        }

        private static EnvDTE80.DTE2 GetDTE2()
        {
            return Package.GetGlobalService(typeof(DTE)) as EnvDTE80.DTE2;
        }
        private string GetSourceFilePath()
        {
            EnvDTE80.DTE2 _applicationObject = GetDTE2();
            UIHierarchy uih = _applicationObject.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            if (null != selectedItems)
            {
                foreach (UIHierarchyItem selItem in selectedItems)
                {
                    ProjectItem prjItem = selItem.Object as ProjectItem;
                    string filePath = prjItem.Properties.Item("FullPath").Value.ToString();
                    //System.Windows.Forms.MessageBox.Show(selItem.Name + filePath);
                    return filePath;
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string currentPath = GetSourceFilePath();
            int first = currentPath.LastIndexOf("\\");
            int second = currentPath.LastIndexOf("\\", first - 1);
            string nameOnly = currentPath.Substring(second + 1, first - second - 1);//.vcxproj
            string folderOnly = currentPath.Substring(0, first + 1);

            string settingsFile = "";
            string[] files = Directory.GetFiles(folderOnly);
            foreach (var file in files)
            {
                if (file.EndsWith(".vcxproj"))
                    settingsFile = file;
                continue;
            }

            var doc = XDocument.Load(settingsFile);
            XNamespace nSpace = "http://schemas.microsoft.com/developer/msbuild/2003";
            var itemElement = nSpace + "ItemDefinitionGroup";
            var elements = doc.Root.Descendants(itemElement);

            foreach (var item in elements)
            {
                string str = (string)item.Attribute("Condition");

                if (str.Contains("x64"))
                {
                    var Link = item.Descendants(nSpace + "Link");

                    foreach (var decendants in Link)
                    {
                        XElement elementToAdd = new XElement(nSpace + "AdditionalLibraryDirectories");
                        elementToAdd.SetValue("$(VCToolsInstallDir)\\lib\\x64;% (AdditionalLibraryDirectories)");
                        decendants.Add(elementToAdd);

                        continue;
                    }

                }

                if (str.Contains("Win32"))
                {
                    var Link = item.Descendants(nSpace + "Link");

                    foreach (var decendants in Link)
                    {
                        XElement elementToAdd = new XElement(nSpace + "AdditionalLibraryDirectories");
                        elementToAdd.SetValue("$(VCToolsInstallDir)\\lib\\x86;% (AdditionalLibraryDirectories)");
                        decendants.Add(elementToAdd);

                        continue;
                    }
                }
            }
            doc.Save(settingsFile);
        }
    }
}
