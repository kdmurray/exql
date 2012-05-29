using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExQL
{
    class Program
    {
        static string _dataFile;
        static ExcelFileType _fileType;
        static ExcelManager mgr;

        static void Main(string[] args)
        {
            ParseParams(args);

            if (_dataFile != null)
            {
                ProcessOpenCommand();
            }

            Console.Write("$==> ");
            string command = Console.ReadLine();

            while (command.ToLower() != "exit")
            {
                string cmdKeyword = "";

                if (command.Contains(" "))
                {
                    cmdKeyword = command.Substring(0, command.IndexOf(" "));
                }
                else
                {
                    cmdKeyword = command;
                }

                switch (cmdKeyword.ToLower())
                {
                    case "select":
                        ProcessSelectCommand(command);
                        break;
                    case "open":
                        ProcessOpenCommand(command);
                        break;
                    default:
                        Console.WriteLine("Command not recognized: {0}", command);
                        break;
                }

                Console.Write("$==> ");
                command = Console.ReadLine();
            }
        }

        private static void ProcessOpenCommand()
        {
            ProcessOpenCommand(_dataFile);
        }

        private static void ProcessOpenCommand(string command)
        {
            string filename = command.Substring(command.IndexOf(" ") + 1);

            if (!File.Exists(filename))
            {
                Console.WriteLine("File '{0}' doesn't exist.", filename);
                Environment.Exit((int)ExitCode.FileNotFound);
            }

            _fileType = ExcelManager.GetFileType(filename);

            if (_fileType == ExcelFileType.INVALID)
            {
                Console.WriteLine("File '{0}' is of an unrecognized file type.", filename);
                Environment.Exit((int)ExitCode.InvalidFileType);
            }

            _dataFile = filename;

            mgr = new ExcelManager(filename);

            Console.WriteLine("\nOPEN: {0}\n", filename);
        }

        private static void ProcessSelectCommand(string command)
        {
            DataTable dt = mgr.Execute(command);

            int[] colWidths = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                colWidths[i] = dt.Columns[i].ColumnName.Length;
            }

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (dr[i].ToString().Length > colWidths[i])
                    {
                        colWidths[i] = dr[i].ToString().Length;
                    }
                }
            }

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                Console.Write(dt.Columns[i].ColumnName.PadRight(colWidths[i]));
                Console.Write(" ");
            }
            Console.WriteLine();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                Console.Write("".PadRight(colWidths[i], '-'));
                Console.Write(" ");
            }
            Console.WriteLine();

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Console.Write(dr[i].ToString().PadRight(colWidths[i]));
                    Console.Write(" ");
                }
                Console.WriteLine("");
            }
        }

        static void ParseParams(string[] args)
        {
            foreach (string param in args)
            {
                if (param.StartsWith("-s=") || param.StartsWith("--source="))
                {
                    _dataFile = param.Substring(param.IndexOf("=") + 1);
                }
            }
        }
    }
}
