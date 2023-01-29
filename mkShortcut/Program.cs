using IWshRuntimeLibrary;
namespace mkShortcut;

class Program
{
    static void Main(string[] args)
    {
        int argscount = 0;
        string name = "";         // -n
        string exe = "";          // -e
        string description = "";  // -c
        string path = "";         // -p
        string hotkey = "";       // -h
        string arguments = "";    // -a
        int window = -1;          // -w
        const string helpMsg = "mkShortcut -n name -e exe_with_path -c comment -p working_path -h hot-key -a arguments -w <Win|Min|Max>";

        while (argscount < args.Length)
        {
            if (args[argscount] == "--help")
            {
                Console.WriteLine("mkShortcut: Help");
                Console.WriteLine(helpMsg);
                return;
            }
            else if (args[argscount] == "-n")
            {
                argscount++;
                name = args[argscount];
            }
            else if (args[argscount] == "-e")
            {
                argscount++;
                exe = args[argscount];
            }
            else if (args[argscount] == "-c")
            {
                argscount++;
                description = args[argscount];
            }
            else if (args[argscount] == "-p")
            {
                argscount++;
                path = args[argscount];
            }
            else if (args[argscount] == "-h")
            {
                argscount++;
                hotkey = args[argscount];
            }
            else if (args[argscount] == "-a")
            {
                argscount++;
                arguments = args[argscount];
            }
            else if (args[argscount] == "-w")
            {
                argscount++;
                switch (args[argscount].ToLower())
                {
                    case "win":
                        window = 1;
                        break;
                    case "min":
                        window = 7;
                        break;
                    case "max":
                        window = 3;
                        break;
                    default:
                        Console.WriteLine("mkShortcut: Error incorrent window specifier.");
                        return;
                }
            }

            argscount++;
        }

        if (name == "")
        {
            Console.WriteLine("mkShortCut: no name given");
            Console.WriteLine(helpMsg);
            return;
        }

        if (exe == "")
        {
            Console.WriteLine("mkShortCut: no exe given");
            Console.WriteLine(helpMsg);
            return;
        }

        object shDesktop = (object)"Desktop";
        WshShell shell = new WshShell();
        string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\"+ name + @".lnk";
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);

        if (description != "")
            shortcut.Description = description;

        if (hotkey != "")
            shortcut.Hotkey = hotkey;

        if (arguments != "")
            shortcut.Arguments = arguments;

        if (path != "")
            shortcut.WorkingDirectory = path;

        if (exe != "")
            shortcut.TargetPath = exe;

        if (window != -1)
            shortcut.WindowStyle = window;
        
        shortcut.Save();

    }
}