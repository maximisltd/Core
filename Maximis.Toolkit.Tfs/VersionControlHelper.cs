using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using System;
using System.Diagnostics;
using System.IO;

namespace Maximis.Toolkit.Tfs
{
    public static class VersionControlHelper
    {
        public static string DownloadAllFiles(VersionControlServer vcs, string tfsPath, string folderPath)
        {
            string rootFolder = null;
            foreach (Item item in vcs.GetItems(tfsPath, RecursionType.Full).Items)
            {
                string itemLocalPath = Path.Combine(folderPath, item.ServerItem.Substring(2).Replace("/", "\\"));
                switch (item.ItemType)
                {
                    case ItemType.Folder:
                        if (Directory.Exists(itemLocalPath)) Directory.Delete(itemLocalPath, true);
                        Directory.CreateDirectory(itemLocalPath);
                        if (string.IsNullOrEmpty(rootFolder)) rootFolder = itemLocalPath;
                        break;

                    case ItemType.File:
                        Trace.WriteLine(string.Format("Downloading file '{0}'", itemLocalPath));
                        item.DownloadFile(itemLocalPath);
                        break;
                }
            }
            return rootFolder;
        }

        public static Workspace GetOrCreateWorkspace(VersionControlServer vcs, string workspaceName, string rootTfsLocation, string rootLocalFolder)
        {
            Workspace ws = vcs.TryGetWorkspace(rootLocalFolder);

            if (ws != null && ws.Name != workspaceName)
            {
                ws.Delete();
                ws = null;
            }

            if (ws == null)
            {
                WorkingFolder workingFolder = new WorkingFolder(rootTfsLocation, rootLocalFolder);
                ws = vcs.CreateWorkspace(new CreateWorkspaceParameters(workspaceName)
                {
                    Folders = new WorkingFolder[] { workingFolder }
                });
            }
            return ws;
        }

        public static VersionControlServer GetVersionControlServer(string projectCollectionUri)
        {
            TfsTeamProjectCollection projectColl = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(projectCollectionUri));
            return projectColl.GetService<VersionControlServer>();
        }
    }
}