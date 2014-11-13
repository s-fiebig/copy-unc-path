using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace CopyUncPath {

    class Program {

        [STAThread]
        static void Main(string[] args) {
            //Takes the parameter from the command line. Since this will
            //be called from the context menu, the context menu will pass it 
            //via %1 (see registry instructions below)
            if (args.Length == 1) {
                Clipboard.SetText(Pathing.GetUNCPath(args[0]));
            } else {
                //This is so you can assign a shortcut to the program and be able to
                //Call it pressing the shortcut you assign. The program will take
                //whatever string is in the Clipboard and convert it to the UNC path
                //For example, you can do "Copy as Path" and then press the shortcut you  
                //assigned to this program. You can then press ctrl-v and it will
                //contain the UNC path
                Clipboard.SetText(Pathing.GetUNCPath(Clipboard.GetText()));
            }
        }
    }

    public static class Pathing {

        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);

        /// <summary>
        /// Given a path, returns the UNC path or the original. (No exceptions
        /// are raised by this function directly). For example, "P:\2008-02-29"
        /// might return: "\\networkserver\Shares\Photos\2008-02-09"
        /// </summary>
        /// <param name="originalPath">The path to convert to a UNC Path</param>
        /// <returns>A UNC path. If a network drive letter is specified, the
        /// drive letter is converted to a UNC or network path. If the 
        /// originalPath cannot be converted, it is returned unchanged.</returns>
        public static string GetUNCPath(string originalPath) {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;

            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':') {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) {
                    int error = WNetGetConnection(originalPath.Substring(0, 2),
                        sb, ref size);
                    if (error == 0) {
                        DirectoryInfo dir = new DirectoryInfo(originalPath);

                        string path = Path.GetFullPath(originalPath)
                            .Substring(Path.GetPathRoot(originalPath).Length);
                        return Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }
            return originalPath;
        }
    }
}

