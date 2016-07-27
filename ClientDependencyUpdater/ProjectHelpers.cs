using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ClientDependencyUpdater
{
    // Borrowed from https://github.com/madskristensen/WebCompiler
    public static class ProjectHelpers
    {
        private static DTE _dte = ClientDependencyUpdaterPackage._dte;

        public static string GetRootFolder(this Project project)
        {
            if (project == null || string.IsNullOrEmpty(project.FullName))
                return null;

            string fullPath;

            try
            {
                fullPath = project.Properties.Item("FullPath").Value as string;
            }
            catch (ArgumentException)
            {
                try
                {
                    // MFC projects don't have FullPath, and there seems to be no way to query existence
                    fullPath = project.Properties.Item("ProjectDirectory").Value as string;
                }
                catch (ArgumentException)
                {
                    // Installer projects have a ProjectPath.
                    fullPath = project.Properties.Item("ProjectPath").Value as string;
                }
            }

            if (string.IsNullOrEmpty(fullPath))
                return File.Exists(project.FullName) ? Path.GetDirectoryName(project.FullName) : null;

            if (Directory.Exists(fullPath))
                return fullPath;

            if (File.Exists(fullPath))
                return Path.GetDirectoryName(fullPath);

            return null;
        }

        public static bool IsKind(this Project project, string kindGuid)
        {
            return project.Kind.Equals(kindGuid, StringComparison.OrdinalIgnoreCase);
        }

        public static IEnumerable<Project> GetAllProjects()
        {
            return _dte.Solution.Projects
                  .Cast<Project>()
                  .SelectMany(GetChildProjects)
                  .Union(_dte.Solution.Projects.Cast<Project>())
                  .Where(p => { try { return !string.IsNullOrEmpty(p.FullName); } catch { return false; } });
        }

        private static IEnumerable<Project> GetChildProjects(Project parent)
        {
            try
            {
                if (!parent.IsKind(ProjectKinds.vsProjectKindSolutionFolder) && parent.Collection == null)  // Unloaded
                    return Enumerable.Empty<Project>();

                if (!string.IsNullOrEmpty(parent.FullName))
                    return new[] { parent };
            }
            catch (COMException)
            {
                return Enumerable.Empty<Project>();
            }

            return parent.ProjectItems
                    .Cast<ProjectItem>()
                    .Where(p => p.SubProject != null)
                    .SelectMany(p => GetChildProjects(p.SubProject));
        }

    }
}
