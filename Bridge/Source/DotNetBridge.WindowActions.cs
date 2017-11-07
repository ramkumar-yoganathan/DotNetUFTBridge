using System.Diagnostics;

namespace DotNetBridge
{
    internal class WindowActions
    {
        public bool HasWindowExists(string title)
        {
            bool hasWindow = false;
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
                if (!string.IsNullOrEmpty(process.MainWindowTitle) && process.MainWindowTitle == title)
                {
                    hasWindow = true;
                    break;
                }
            }
            return hasWindow;
        }
    }
}
