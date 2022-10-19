//-----------------------------------------------------------------------
// The Extractor V1.1.4
//-----------------------------------------------------------------------

using System.Text;

[assembly: System.CLSCompliant(true)]
namespace Microsoft.Windows.Kits.Samples
{
    using System;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Linq;
    using System.Collections.Generic;
    using System.Xml;
    //The below dependencies are dll files that are part of HLK Studio
    //You need to manually add the dll files as reference to this project
    using Microsoft.Windows.Kits.Hardware.FilterEngine;
    using Microsoft.Windows.Kits.Hardware.ObjectModel;
    using Microsoft.Windows.Kits.Hardware.ObjectModel.DBConnection;
    using Microsoft.Windows.Kits.Hardware.ObjectModel.Submission;


    internal static class ProgramSettings
    {
        internal static string PackagesDir = null;
        internal static string PackageFile = null;
        internal static bool ExtractLogs = false;
        internal static string LogsDir = null;
        internal static TextWriter Log = null;
        internal static string LogFile = null;
        internal static List<string> DevelopmentPhases = null;
    }

    internal static class Constants
    {
        internal const string PackageExt = ".hlkx";
        internal const string DefaultLogName = "PackageAnalysisLog.txt";
    }

    public class PackageLogExtractor
    {
        static void Main(string[] args)
        {
            if (false == ParseArgs(args))
            {
                ShowUsage();
                return;
            }
            PackageAnalyze(); // all the command line work.

            if (ProgramSettings.Log != null)
            {
                ProgramSettings.Log.Dispose();
            }
        }
        static void ShowUsage()
        {
            string usage = "";


            usage += Environment.NewLine + "=============================HLK Filter==============================="+
                     Environment.NewLine + "Version: 1.1.3  "+
                     Environment.NewLine + "Release Date: MAR-12-2020"+
                     Environment.NewLine + "Supported HLK Versions: 19H1, 19H2"+
                     Environment.NewLine + "======================================================================" +
                     Environment.NewLine + "Usage: "+
                     Environment.NewLine + "HLK_Log_Extractor.exe [/PackagesDir=<path>]" +
                     Environment.NewLine + "                   [/PackageFile=<path>]" +
                     Environment.NewLine +
                     Environment.NewLine + "Any parameter in [] is optional." +
                     Environment.NewLine + "At least /PackagesDir or /PackageFile must be specified" +
                     Environment.NewLine +
                     Environment.NewLine + @"HLK_Log_Extractor.exe /PackageFile=C:\Users\user\Desktop\file.hlkx" +
                     Environment.NewLine + @"HLK_Log_Extractor.exe /PackagesDir=C:\Users\user\Desktop\Hlk\" +
                     Environment.NewLine + "======================================================================" +
                     Environment.NewLine + "Parameter Descriptions:" +
                     Environment.NewLine + "======================================================================" +
                     Environment.NewLine +
                     Environment.NewLine + "PackagesDir:   Directory to recursively search for package files." +
                     Environment.NewLine +
                     Environment.NewLine + "PackageFile:   Path to single package file to process (.hlkx or .hlkp)" +
                     Environment.NewLine +
                     Environment.NewLine + "Logs will be extracted in the current directory" +
                     Environment.NewLine;
            //*** Need to update that

            Console.WriteLine(usage);

        }


        static bool ParseArgs(string[] args)
        {

            if ((args.Length == 0) || (args[0].Contains("?")))
            {
                return false;
            }

            foreach (string arg in args)
            {
                if (arg.StartsWith("/PackagesDir=", StringComparison.OrdinalIgnoreCase))
                {
                    ProgramSettings.PackagesDir = Path.GetFullPath(arg.Substring("/PackagesDir=".Length));
                }
                else if (arg.StartsWith("/PackageFile=", StringComparison.OrdinalIgnoreCase))
                {
                    ProgramSettings.PackageFile = Path.GetFullPath(arg.Substring("/PackageFile=".Length));
                }
                else
                {
                    Console.WriteLine("Unrecognized parameter: " + arg);
                    return false;
                }
            }
            return true;
        }

        private static void WriteMessage(string message, params object[] args)
        {
            if (ProgramSettings.Log != null)
            {
                if (args.Length > 0)
                {
                    ProgramSettings.Log.WriteLine(message, args);
                }
                else
                {
                    ProgramSettings.Log.WriteLine(message);
                }
            }
            else
            {
                Console.WriteLine(message, args);
            }
        }

        /// 
        /// <summary>
        /// The Parse engine.
        /// </summary>
        /// 
        public static void PackageAnalyze()
        {
            string[] packageFiles;
            PackageManager manager = null;
            List<DevelopmentPhase> developmentPhases = null;

            //
            // Check development phases
            //
            var phases = (DevelopmentPhase[])Enum.GetValues(typeof(DevelopmentPhase)); 
            developmentPhases = new List<DevelopmentPhase>(phases);
            
            if (string.IsNullOrEmpty(ProgramSettings.PackageFile) == false)
            {
                packageFiles = new string[] { ProgramSettings.PackageFile };
            }
            else
            {
                //*** Same extension?
                packageFiles = Directory.GetFiles(ProgramSettings.PackagesDir, "*.hlk?", SearchOption.AllDirectories);
            }

            if (packageFiles.Count() == 0)
            {
                Console.WriteLine("No *.HLK files in the specified directory");
            }
            else
            {
                foreach (var file in packageFiles)
                {
                    try
                    {
                        ProcessFile(manager, file, developmentPhases);
                    }
                    catch (Exception e)
                    {
                        WriteMessage(e.ToString());
                        Console.WriteLine(e);
                        return;
                    }
                    finally
                    {
                        if (manager != null)
                        {
                            manager.Dispose();
                        }
                        manager = null;
                    }

                    if (ProgramSettings.Log != null)
                    {
                        ProgramSettings.Log.Flush();
                    }
                }
            }
        }

        static void ProcessFile(ProjectManager manager, string file, List<DevelopmentPhase> developmentPhases)
        {
            Console.WriteLine("==============================================================================");
            WriteMessage("Process package {0}.", file);
            manager = new PackageManager(file);

            List<Project> packageProjects = new List<Project>();
            ReadOnlyCollection<string> projectNames = manager.GetProjectNames();

            foreach (var projectName in projectNames)
            {
                WriteMessage("Validating project: " + projectName);

                bool packageHasResults = false;
                ProjectInfo info = manager.GetProjectInfo(projectName);

                //
                // There wasn't anything run, so no files are going to exist.
                // The OM is going to choke if you try to extract the log files later
                // on.
                //
                if (info.TotalCount == 0)
                {
                    WriteMessage("No results.");
                    continue;
                }
                else
                {
                    WriteMessage(
                        string.Format(
                        ("Results count: \n\t" +
                        "Pass - {0} Fail - {1} Running - {2} NotRun - {3} Total - {4}"),
                        info.PassedCount,
                        info.FailedCount,
                        info.RunningCount,
                        info.NotRunCount,
                        info.TotalCount));

                    packageHasResults = true;
                }

                if (packageHasResults)
                {
                    var project = manager.GetProject(projectName);
                    packageProjects.Add(project);
                    IList<Test> testList = project.GetTests().SelectMany((t => t.DevelopmentPhases.Join(developmentPhases, p => p, d => d, (p, d) => t))).ToList();


                    ReadOnlyCollection<IFilter> appliedFilters = project.GetAppliedFilters();
                    Console.Out.WriteLine("{0} filters were applied for project {1}", appliedFilters.Count, projectName);

                    var csv = new StringBuilder();
                    var newLine = string.Format("Project Name, Project Manager Version, Filter Number, Expiration Date, Version, Status, Type, Title, Description, Resolution, Last Modified Date");
                    csv.AppendLine(newLine);
                    foreach (var filter in appliedFilters)
                    {
                        var zero = project.Name;
                        var os_ver = project.ProjectManager.Version;
                        var first = filter.FilterNumber;
                        var second = filter.ExpirationDate;
                        var third = filter.Version;
                        var fourth = filter.Status;
                        var fifth = filter.Type;
                        var sixth = filter.Title.ToString().Replace(',', ' ').Replace('\n', ' ').Trim();
                        var seventh = filter.IssueDescription.ToString().Replace(',', ' ').Replace('\n', ' ').Trim();
                        var eigth = filter.IssueResolution.ToString().Replace(',', ' ').Replace('\n', ' ').Trim();
                        var ninth = filter.LastModifiedDate;

                        var newLine2 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", zero, os_ver, first, second, third, fourth, fifth, sixth, seventh, eigth, ninth);
                        csv.AppendLine(newLine2);

                    }

                    String DirPathLog = AppDomain.CurrentDomain.BaseDirectory + project.Name;
                    DirectoryInfo di = Directory.CreateDirectory(DirPathLog);

                    try
                    {
                        File.WriteAllText(DirPathLog + @"\" + project.Name + "_Filters.csv", csv.ToString());
                    }
                    catch (Exception e)
                    {
                        Console.BackgroundColor = ConsoleColor.White;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Unable to write the filter file. Check if file already exists and is open or check access.");
                        WriteMessage(e.ToString());
                        Console.WriteLine(e);
                        throw;
                    }
                    
                    var csv2 = new StringBuilder();
                    newLine = string.Format("Project Name, Project Manager Version, Filter Applied, Test ID, Test Name, Description, Test Status");
                    csv2.AppendLine(newLine);

                    string DuplicateChecker = null;

                    foreach (Test test in testList.Where(t => (t.GetTestResults().Count > 0)))
                    {
                        var zero = projectName;
                        var os_ver = project.ProjectManager.Version;
                        var test_ID = test.Id;
                        var first = test.AreFiltersApplied;
                        var second = test.Name.ToString().Replace(',', ' ').Replace('\n', ' ').Replace('\r', ' ').Trim();
                        var third = test.Description.ToString().Replace(',', ' ').Replace('\n', ' ').Replace('\r', ' ').Trim();
                        var fourth = test.Status;

                        newLine = string.Format("{0},{1},{2},{3},{4},{5},{6}", zero, os_ver, first, test_ID, second, third, fourth);

                        if (second != DuplicateChecker)
                        {
                            csv2.AppendLine(newLine);
                            DuplicateChecker = second;
                        }
                    }

                    Console.WriteLine("OS Version: "+ project.ProjectManager.Version);

                    try
                    {
                        File.WriteAllText(DirPathLog + @"\" + project.Name + "_Tests.csv", csv2.ToString());
                    }
                    catch (Exception e)
                    {
                        Console.BackgroundColor = ConsoleColor.White;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Unable to write the filter file. Check if file already exists and is open or check access.");
                        WriteMessage(e.ToString());
                        Console.WriteLine(e);
                        throw;
                    }
                    
                }

            }
        }

    }
}