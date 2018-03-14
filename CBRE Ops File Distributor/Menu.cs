using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;



namespace CBRE_Ops_File_Distributor {
    public class Menu {

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int cmdShow);
                
        static Menu() {
            Maximize();
        }

        private static void Maximize() {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3); //SW_MAXIMIZE = 3
        }


        public static void Header() {
            Console.WriteLine("Program:       CBRE NZ Ops File Distributor");
            Console.WriteLine("Developer:     Todd Sandford");
            Console.WriteLine("Date:          12 February 2018");
            Console.WriteLine("Updated:       8 March 2018");
            Console.WriteLine("Update Notes:  Added ");
            Console.WriteLine("\n******************************************************************************************************\n");
        }



        public static void CheckPassword() {
            Console.Write("Please enter CBRE NZ Ops password: ");

            string pass = "";
            ConsoleKeyInfo key;

            do {
                key = Console.ReadKey(true);
                
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter) {
                    pass += key.KeyChar;
                    Console.Write("*");
                } else {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0) {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            } while (key.Key != ConsoleKey.Enter); // stop receiving input once enter is pressed

            if(pass != "Ultron") {
                Exit();
            }
            Console.WriteLine("\n");
        }

        public static bool Continue() {
            string ans = "";
            bool validResponse = false;

            while (!validResponse) {
                Console.Write("Would you like to continue? (y/n): ");
                ans = Console.ReadLine();
                if (ans.ToLower() == "y" || ans.ToLower() == "n") {
                    validResponse = true;
                } else {
                    Console.WriteLine("Please enter a valid response (y or n).");
                }
            }

            if (ans.ToLower() == "n") {
                Exit();
            }

            return true;

        }
        
        private static void Exit() {
            Console.Write("\nPress any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static void Finish() {
            Console.WriteLine("\nFinished.");
            Exit();
        }
    }
}
