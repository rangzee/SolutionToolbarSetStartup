//------------------------------------------------------------------------------
// <copyright file="SetStartupButton.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SolutionToolbarSetStartup
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SetStartupButton
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandStartupId = 0x0100;
        public const int CommandUnloadId = 0x0101;
        public const int CommandReloadId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("977f99b5-05cb-4e0f-9a3c-31500f5659ac");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetStartupButton"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private SetStartupButton(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuStartupCommandID = new CommandID(CommandSet, CommandStartupId);
                var menuStartupItem = new OleMenuCommand(this.StartupMenuItemCallback, menuStartupCommandID);
                menuStartupItem.BeforeQueryStatus += menuStartupItem_BeforeQueryStatus;
                commandService.AddCommand(menuStartupItem);

                var menuUnloadCommandID = new CommandID(CommandSet, CommandUnloadId);
                var menuUnloadItem = new OleMenuCommand(this.UnloadMenuItemCallback, menuUnloadCommandID);
                menuUnloadItem.BeforeQueryStatus += menuUnloadItem_BeforeQueryStatus;
                commandService.AddCommand(menuUnloadItem);

                var menuReloadCommandID = new CommandID(CommandSet, CommandReloadId);
                var menuReloadItem = new OleMenuCommand(this.ReloadMenuItemCallback, menuReloadCommandID);
                menuReloadItem.BeforeQueryStatus += menuReloadItem_BeforeQueryStatus;
                commandService.AddCommand(menuReloadItem);
            }
        }

        public void menuStartupItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                menuCommand.Visible = false;
                Array projects = GetActiveSolutionProjects(sender);
                if (projects.Length == 1
                    && string.Compare(EnvDTE.Constants.vsProjectKindUnmodeled, ((EnvDTE.Project)projects.GetValue(0)).Kind, StringComparison.OrdinalIgnoreCase) != 0
                    && string.Compare(EnvDTE.Constants.vsProjectKindSolutionItems, ((EnvDTE.Project)projects.GetValue(0)).Kind, StringComparison.OrdinalIgnoreCase) != 0
                    )
                {
                    menuCommand.Visible = true;
                }
            }
        }

        public void menuUnloadItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                menuCommand.Visible = false;
                Array projects = GetActiveSolutionProjects(sender);
                foreach (EnvDTE.Project project in projects)
                {
                    //if (project.FullName.EndsWith(".csproj"))
                    if (string.Compare(EnvDTE.Constants.vsProjectKindUnmodeled, project.Kind, StringComparison.OrdinalIgnoreCase) != 0
                        || string.Compare(EnvDTE.Constants.vsProjectKindSolutionItems, project.Kind, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        menuCommand.Visible = true;
                        break;
                    }
                }
            }
        }

        public void menuReloadItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                menuCommand.Visible = false;
                Array projects = GetActiveSolutionProjects(sender);
                if (projects.Length == 1 && string.Compare(EnvDTE.Constants.vsProjectKindSolutionItems, ((EnvDTE.Project)projects.GetValue(0)).Kind, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    menuCommand.Visible = true;
                }
                if (projects.Length > 1)
                {
                    foreach (EnvDTE.Project project in projects)
                    {
                        //if (project.FullName.EndsWith(".csproj"))
                        if (string.Compare(EnvDTE.Constants.vsProjectKindUnmodeled, project.Kind, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            menuCommand.Visible = true;
                            break;
                        }
                    }
                }
            }
        }

        private Array GetActiveSolutionProjects(object sender)
        {
            EnvDTE.DTE dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));

            if (dte != null)
            {
                var projects = (Array)dte.ActiveSolutionProjects;
                if (projects.Length > 0)
                {
                    var first = projects.GetValue(0);
                    var project = (EnvDTE.Project)first;
                    var kind = project.Kind;
                    System.Diagnostics.Debug.WriteLine(kind);
                }
            }

            return dte == null ? null : (Array)dte.ActiveSolutionProjects;
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SetStartupButton Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new SetStartupButton(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void StartupMenuItemCallback(object sender, EventArgs e)
        {
            // Ref: https://msdn.microsoft.com/en-us/library/bb166773.aspx?f=255&MSPPError=-2147217396

            EnvDTE.DTE dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
            if (dte == null)
            {
                throw new Exception("Unable to retrieve DTE");
            }

            //Array activeSolutionProjects = (Array)dte.ActiveSolutionProjects;
            //if (activeSolutionProjects == null || activeSolutionProjects.Length == 0)
            //{
            //    throw new Exception("Active Solution Projects is null");
            //}

            //EnvDTE.Project dteProject = (EnvDTE.Project)activeSolutionProjects.GetValue(0);
            //if (dteProject == null)
            //{
            //    throw new Exception("Active Solution Projects [0] is null");
            //}

            try
            {
                dte.ExecuteCommand("Project.SetasStartUpProject");
            }
            catch { }

            //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "SetStartupButton";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private void UnloadMenuItemCallback(object sender, EventArgs e)
        {
            EnvDTE.DTE dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
            if (dte == null)
            {
                throw new Exception("Unable to retrieve DTE");
            }

            try
            {
                dte.ExecuteCommand("Project.UnloadProject");
            }
            catch { }
        }

        private void ReloadMenuItemCallback(object sender, EventArgs e)
        {
            EnvDTE.DTE dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
            if (dte == null)
            {
                throw new Exception("Unable to retrieve DTE");
            }

            try
            {
                dte.ExecuteCommand("Project.ReloadProject");
            }
            catch { }
        }
    }
}
