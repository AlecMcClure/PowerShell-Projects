using System.Diagnostics;
using System.IO;

namespace explorer
{
    public class program 
    {
        public void Main()
        {
            openFolder(@"C:\Users\");
        }
        private void openFolder(string folderPath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                Arguments = folderPath,
                FileName = "explorer.exe"
            };

            Process.Start(startInfo);
        }
    }
}