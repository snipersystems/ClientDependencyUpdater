//------------------------------------------------------------------------------
// <copyright file="VSPackage1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Linq;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE;
using EnvDTE80;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.Collections.Concurrent;
using System.Threading;

namespace ClientDependencyUpdater
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.1", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(ClientDependencyUpdaterPackage.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ClientDependencyUpdaterPackage : Package
    {
        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "f3866fb2-3349-4063-a81c-d9fe12836290";

        public static DTE _dte;
        private Events _dteEvents;
        //private DocumentEvents _documentEvents;
        private SolutionEvents _solutionEvents;

        private IVsStatusbar _statusBar = null;
        public IVsStatusbar StatusBar
        {
            get
            {
                if (_statusBar == null)
                {
                    _statusBar = GetService(typeof(SVsStatusbar)) as IVsStatusbar;
                }

                return _statusBar;
            }
        }

        private ConcurrentDictionary<Project, FileSystemWatcher> _watchers;



        /// <summary>
        /// Initializes a new instance of the <see cref="ClientDependencyUpdaterPackage"/> class.
        /// </summary>
        public ClientDependencyUpdaterPackage()
        {
            _watchers = new ConcurrentDictionary<Project, FileSystemWatcher>();
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            _dte = (DTE)GetService(typeof(SDTE));
            _dteEvents = _dte.Events;
            //_documentEvents = _dteEvents.DocumentEvents;
            //_documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;

            _solutionEvents = _dteEvents.SolutionEvents;

            _solutionEvents.Opened += Opened;
            _solutionEvents.AfterClosing += AfterClosing;
            _solutionEvents.ProjectAdded += WatchProject;
            _solutionEvents.ProjectRemoved += RemoveProject;

        }

        private void AfterClosing()
        {
            for (int i = _watchers.Count - 1; i >= 0; i--)
            {
                RemoveProject(_watchers.ElementAt(i).Key);
            }
        }

        private void Opened()
        {
            var projects = ProjectHelpers.GetAllProjects();
            foreach (var project in projects)
            {
                WatchProject(project);
            }
        }

        private void WatchProject(Project project)
        {
            if (_watchers.ContainsKey(project))
                return;

            var fsw = new FileSystemWatcher(project.GetRootFolder());
            fsw.Created += FileChanged;
            fsw.Changed += FileChanged;
            fsw.Renamed += FileChanged;
            fsw.Deleted += FileChanged;

            fsw.IncludeSubdirectories = true;
            fsw.NotifyFilter = NotifyFilters.Size | NotifyFilters.CreationTime | NotifyFilters.FileName;
            fsw.EnableRaisingEvents = true;

            _watchers.TryAdd(project, fsw);
        }

        private void RemoveProject(Project project)
        {
            if (project == null || !_watchers.ContainsKey(project))
                return;

            FileSystemWatcher fsw;
            _watchers[project].Dispose();
            _watchers.TryRemove(project, out fsw);
        }

        private void FileChanged(object sender, FileSystemEventArgs e)
        {
            string filename = Path.GetFileName(e.FullPath);

            if (filename.Contains("~"))
                return; // VS adds ~ to temp filenames

            string ext = Path.GetExtension(filename);

            if (ext != ".js" && ext != ".css")
                return;

            var fsw = (FileSystemWatcher)sender;
            var project = _watchers.Keys.FirstOrDefault(p => e.FullPath.StartsWith(p.GetRootFolder()));

            UpdateClientDependencyVersion(project);
        }

        private void UpdateClientDependencyVersion(Project project)
        {
            uint cookie = 0;
            string version = Stopwatch.GetTimestamp().ToString().Substring(5);

            var configFiles = GetConfigFiles(project.ProjectItems).ToArray();
            for (uint i = 0; i < configFiles.Length; i++)
            {
                var configFile = configFiles[i];
                StatusBar.Progress(ref cookie, 1, "Updating Client Dependency Version", i + 1, (uint)configFiles.Length + 1);

                Window openConfigFile = null;
                if (!configFile.IsOpen)
                {
                    openConfigFile = configFile.Open();
                }
                string input = GetContent(configFile.Document);
                string output;
                if (UpdateVersion(version, input, out output))
                {
                    SetContent(configFile.Document, output);
                }
                if (openConfigFile != null)
                    openConfigFile.Close(vsSaveChanges.vsSaveChangesYes);
            }
            StatusBar.Progress(ref cookie, 0, "", 0, 0);
            StatusBar.FreezeOutput(0);
            StatusBar.Clear();
        }

        private static bool UpdateVersion(string version, string content, out string output)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(content);

            var clientDependencyNodes = xmlDocument.SelectNodes("//clientDependency");
            foreach (XmlElement node in clientDependencyNodes)
            {
                if (!node.HasAttribute("configSource"))
                    node.SetAttribute("version", version);
            }
            if (clientDependencyNodes.Count == 0)
            {
                output = null;
                return false;
            }
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "\t",
                NewLineChars = "\r\n",
                NewLineHandling = NewLineHandling.Replace
            };

            using (var sw = new StringWriter())
            using (var xw = XmlWriter.Create(sw, settings))
            {
                xmlDocument.WriteTo(xw);
                xw.Flush();
                content = sw.GetStringBuilder().ToString();
            }

            output = content;
            return true;
        }

        private static string GetContent(Document document)
        {
            var textDocument = (TextDocument)document.Object(nameof(TextDocument));
            EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
            string content = editPoint.GetText(textDocument.EndPoint);
            return content;
        }

        private static void SetContent(Document document, string content)
        {
            var textDocument = (TextDocument)document.Object(nameof(TextDocument));
            EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
            EditPoint endPoint = textDocument.EndPoint.CreateEditPoint();
            editPoint.ReplaceText(endPoint, content, 0);
            document.Save();
        }

        private IEnumerable<ProjectItem> GetConfigFiles(ProjectItems projectItems)
        {
            foreach (ProjectItem projectItem in projectItems)
            {
                if (projectItem.ProjectItems.Count > 0)
                {
                    foreach (var childItem in GetConfigFiles(projectItem.ProjectItems))
                    {
                        yield return childItem;
                    }
                }

                for (short i = 0; i < projectItem.FileCount; i++)
                {
                    string filename = Path.GetFileName(projectItem.FileNames[i]).ToLower();

                    if (filename == "web.config" || filename == "clientdependency.config")
                        yield return projectItem;
                }

            }
        }
        #endregion
    }
}
