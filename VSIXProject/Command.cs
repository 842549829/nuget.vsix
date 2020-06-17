using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject
{
    /*
     * 官网参考地址:https://docs.microsoft.com/zh-cn/visualstudio/extensibility/starting-to-develop-visual-studio-extensions?view=vs-2019
     * dll 打包 https://www.cnblogs.com/mq0036/p/12737309.html
     *
     * 添加菜单 https://www.cnblogs.com/sdwdjzhy/p/7300480.html
     */

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        private static DTE2 _dte;

        /// <summary>
        /// DotnetPublishRelease.
        /// </summary>
        public const int DotnetPublishRelease = 0x0100;

        public const int DotnetPublishDebug = 0x0101;

        public const int NugetPushRelease = 0x0102;

        public const int NugetPushDebug = 0x0103;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("be015ebc-4ec5-4a91-9add-8223d8a7b003");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            commandService.AddCommand(new MenuCommand(DotnetPublishReleaseExecute, new CommandID(CommandSet, DotnetPublishRelease)));
            commandService.AddCommand(new MenuCommand(DotnetPublishDebugExecute, new CommandID(CommandSet, DotnetPublishDebug)));
            commandService.AddCommand(new MenuCommand(NugetPushReleaseExecute, new CommandID(CommandSet, NugetPushRelease)));
            commandService.AddCommand(new MenuCommand(NugetPushDebugExecute, new CommandID(CommandSet, NugetPushDebug)));
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
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
            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;

            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command(package, commandService);

            _dte = dte;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void DotnetPublishReleaseExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            #region 执行项目

            /*
            var solution = _dte.Solution;
            //var projects = solution.Projects;

            var projects = (UIHierarchyItem[])_dte?.ToolWindows.SolutionExplorer.SelectedItems;
            var project = projects[0].Object as Project;

            //获取项目所有引用
            //var vsproject = project.Object as VSLangProj.VSProject;
            //foreach (VSLangProj.Reference reference in vsproject.References)
            //{

            //}

            var solutionName = Path.GetFileName(solution.FullName);
            var solutionDir = Path.GetDirectoryName(solution.FullName);
            var projectName = Path.GetFileName(project.FullName);
            var projectDir = Path.GetDirectoryName(project.FullName); 
            */


            // 创建cmd
            #endregion

            DotentPublish("Release");
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void DotnetPublishDebugExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DotentPublish("Debug");
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void NugetPushReleaseExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            NugetPush("Release");
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void NugetPushDebugExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            NugetPush("Debug");
        }

        private void DotentPublish(string configuration)
        {
            System.Diagnostics.Process p = null;
            try
            {
                var projects = (UIHierarchyItem[])_dte?.ToolWindows.SolutionExplorer.SelectedItems;
                var project = projects[0].Object as Project;
                var projectDir = Path.GetDirectoryName(project.FullName);
                var projectPath = Path.Combine(projectDir ?? throw new InvalidOperationException(), $"bin\\{configuration}");
                p = new System.Diagnostics.Process();
                p.StartInfo.FileName = @"C:\WINDOWS\system32\cmd.exe ";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                p.StandardInput.WriteLine(projectDir.Substring(0, 2));
                p.StandardInput.WriteLine($"del {projectPath}\\*.nupkg /s/q");  //先转到系统盘下
                p.StandardInput.WriteLine($"dotnet pack -s {Path.Combine(projectDir, Path.GetFileName(project.FullName))} -c {configuration}");  //先转到系统盘下
                p.StandardInput.WriteLine(@"cd C:\WINDOWS\system32 ");  //再转到CMD所在目录下
                p.StandardInput.WriteLine($"dotnet nuget push {projectPath}\\*.nupkg -k Benchint -s http://10.3.1.240:8080/nuget");
                p.StandardInput.WriteLine("exit");
                p.WaitForExit();
                ShowMessageBox("success");
            }
            catch (Exception ex)
            {
                ShowMessageBox(ex.Message);
            }
            finally
            {
                p?.Close();
                p?.Dispose();
            }
        }

        private void NugetPush(string configuration)
        {
            System.Diagnostics.Process p = null;
            try
            {
                var projects = (UIHierarchyItem[])_dte?.ToolWindows.SolutionExplorer.SelectedItems;
                var project = projects[0].Object as Project;
                var projectDir = Path.GetDirectoryName(project.FullName);
                var projectPath = Path.Combine(projectDir ?? throw new InvalidOperationException());
                p = new System.Diagnostics.Process();
                p.StartInfo.FileName = @"C:\WINDOWS\system32\cmd.exe ";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                p.StandardInput.WriteLine(projectDir.Substring(0, 2));
                p.StandardInput.WriteLine($"cd {projectDir.Substring(2)}");
                p.StandardInput.WriteLine($"del {projectPath}\\*.nupkg /s/q");  //先转到系统盘下
                p.StandardInput.WriteLine($"nuget pack {project.Name}.csproj -Prop Configuration={configuration}");
                p.StandardInput.WriteLine($"nuget push *.nupkg -Source http://10.3.1.240:8080/nuget -ApiKey Benchint");
                p.StandardInput.WriteLine("exit");
                p.WaitForExit();
                ShowMessageBox("success");
            }
            catch (Exception ex)
            {
                ShowMessageBox(ex.Message);
            }
            finally
            {
                p?.Close();
                p?.Dispose();
            }
        }

        private void ShowMessageBox(string message)
        {
            string title = "Command";
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}