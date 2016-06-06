using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Diagnostics;

namespace ReflectWILinks
{
    class Program
    {
        private static ConsoleColor defaultConsoleColor;

        static void Main(string[] args)
        {
            Console.WriteLine("ReflectWILinks by Vincent Labatut (2012) - Visual Studio ALM MVP");
            Console.WriteLine("A utility to restore work items links after a TFS to TFS migration");
            Console.WriteLine();

            Stopwatch watch = new Stopwatch();
            watch.Start();
            defaultConsoleColor = Console.ForegroundColor;

            var reflector = new WiLinksReflector(ProcessParameters.Default.SourceTfsUri, ProcessParameters.Default.TargetTfsUri, ProcessParameters.Default.TargetProject);
            reflector.LogMessageEvent += refl_LogMessageEvent;
            reflector.AddMissingRelatedLinks = ProcessParameters.Default.AddMissingRelated;
            reflector.AddMissingChangesetsLinks = ProcessParameters.Default.AddMissingChangesets;
            reflector.AddMissingExternalLinks = ProcessParameters.Default.AddMissingOtherExternal;

            reflector.LoadReflectedWorkItemIds(ProcessParameters.Default.ScopeQueryGuid);

            reflector.ProcessWorkItems(ProcessParameters.Default.TargetQueryGuid);

            Console.ForegroundColor = defaultConsoleColor;
            watch.Stop();
            Console.WriteLine();

            Console.WriteLine("Processed workitems: " + reflector.NbProcessedWorkItems);
            Console.WriteLine("Elapsed time: " + watch.Elapsed);
            Console.WriteLine("Save error(s): " + reflector.NbSaveErrors);
            Console.WriteLine("Missing target workitems for related links: " + reflector.NbMissingRelatedWorkItems);
            Console.WriteLine("Related links: found {0} - existing {1} - crosslinks {2} - added {3}",
                reflector.NbSourceRelatedLinksFound,
                reflector.NbTargetRelatedLinksFound,
                reflector.NbTargetCrossRelatedLinks,
                reflector.NbTargetRelatedLinksAdded);
            Console.WriteLine("Changeset links: found {0} - existing {1} - added {2}",
                reflector.NbSourceChangesetLinksFound,
                reflector.NbTargetChangesetLinksFound,
                reflector.NbTargetChangesetLinksAdded);
            Console.WriteLine("External links: found {0} - existing {1} - added {2}",
                reflector.NbSourceExternalLinksFound,
                reflector.NbTargetExternalLinksFound,
                reflector.NbTargetExternalLinksAdded);
        }

        static void refl_LogMessageEvent(string message, TraceLevel level)
        {
            switch (level)
            {
                case TraceLevel.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case TraceLevel.Off:
                case TraceLevel.Info:
                    Console.ForegroundColor = defaultConsoleColor;
                    break;
                case TraceLevel.Verbose:
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    break;
                case TraceLevel.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                default:
                    Console.ForegroundColor = defaultConsoleColor;
                    break;
            }
            Console.WriteLine(message);
        }
    }
}
