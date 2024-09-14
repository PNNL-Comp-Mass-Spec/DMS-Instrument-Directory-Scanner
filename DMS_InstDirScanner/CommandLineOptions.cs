using PRISM;

namespace DMS_InstDirScanner
{
    internal class CommandLineOptions
    {
        // Ignore Spelling: bionet, DMS

        [Option("noBionet", HelpShowsDefault = false, HelpText = "Skip instruments on bionet")]
        public bool NoBionet { get; set; }

        [Option("preview", HelpShowsDefault = false, HelpText = "Search for files and directories, but do not update any files on the DMS_InstSourceDirScans share")]
        public bool PreviewMode { get; set; }
    }
}
