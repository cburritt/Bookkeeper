/*
 * Author: Caden Burritt
 */


using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Bookkeeper
{
    public class Model
    {
        public static Control control = new Control();

        [STAThread]
        public static void Main(string[] args)
        {
            

            while (true)
            {
                Console.WriteLine("Welcome To BookKeeping Buddy");
                Console.WriteLine("What would you like to do");
                Console.WriteLine("1) Add receipt to exel by picture");
                Console.WriteLine("2) Add receipt Manually");
                Console.WriteLine("3) Change exel file path");
                Console.WriteLine("4) Sort Exel sheet");
                Console.WriteLine("5) Exit Program");
                Console.Write("1/2/3: ");
                string input = Console.ReadLine();
                Console.WriteLine("\n\n\n\n\n\n\n\n");

                if (input.Equals("5"))
                {
                    Environment.Exit(0);
                }
                else if (input.Equals("1"))
                {
                    AddReceipt();
                }
                else if (input.Equals("2"))
                {
                    AddReceiptManual();

                }
                else if (input.Equals("4")) 
                {
                    sort();
                }







            }
        }
        public static void  AddReceipt()
        {
            string filePath;
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder";
                dialog.UseDescriptionForTitle = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Console.WriteLine("You selected folder: " + dialog.SelectedPath);
                    filePath = dialog.SelectedPath;
                }
                else
                {
                    Console.WriteLine("No folder selected.");
                    return;
                }
            }
            if (Directory.Exists(filePath))
            {
                int fileCount = Directory.GetFiles(filePath ).Length;
                Console.WriteLine("Number of files: " + fileCount);
            }
            else
            {
                Console.WriteLine("Folder does not exist.");
            }
            string[] receipts = Directory.GetFiles(filePath, "*.*")
            .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                       file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                       file.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                       file.EndsWith(".gif", StringComparison.OrdinalIgnoreCase) ||
                       file.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase))
            .ToArray();

            
            for (int i = 0; i < receipts.Length; i++) 
            {
                Console.WriteLine(i + 1 + ": " + receipts[i]);
                string total;
                string date;
                string dept;
                string description;
                while (true) 
                {
                    date = control.ReadReceiptDate(receipts[i]);
                    total = control.ReadReceiptTotal(receipts[i]);
                    
                    Console.WriteLine("Date: " + date + "\n Total:" + total + "\nCorrect?");
                    Console.Write("Y/N: ");
                    string input = Console.ReadLine();
                    if (total == "No total found")
                    {
                        input = "N";
                    }
                    if (!input.ToUpper().Equals("Y"))
                    {
                        Console.Write("Enter the Date: ");
                        date = Console.ReadLine();
                        Console.Write("Enter the Total: ");
                        total = Console.ReadLine();
                    }
                    Console.Write("Enter Description: ");
                    description = Console.ReadLine();
                    Console.Write("Enter Department: ");
                    dept = Console.ReadLine();
                    if(date != null && total != null && dept != null && description != null)
                    {
                        Console.WriteLine("Valid Entry");
                        control.AddReceipt(dept, date, description, float.Parse(total));
                        Console.WriteLine("Adding Receipt");
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Invalid Entry.....");
                    }
                }
                
            }
            Console.WriteLine("Press enter to Push receipts to the exel....");
            string x = Console.ReadLine();
            if (x == "")
            {
                Console.WriteLine("Pushing to exel.....");

                for(int i = 0; i < receipts.Length; i++ )
                {
                    control.WriteToExel(control.receipts[i]);
                    Console.WriteLine("Added: " + receipts[i]);
                    
                }
                control.receipts.Clear();
            }
            Console.WriteLine("\n\n\n\n\n\n\n\n");
        }

        public static void AddReceiptManual()
        {
            string date, dept, description;
            float total;
            string pattern = @"^(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/\d{4}$";

            while (true)
            {
                Console.WriteLine("Manual Receipt Input");

                // validate date format
                while (true)
                {
                    Console.Write("Input the Date (mm/dd/yyyy): ");
                    date = Console.ReadLine();
                    if (Regex.IsMatch(date, pattern))
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Incorrect Formatting");
                    }
                }

                Console.Write("Input the Total: ");
                total = float.Parse(Console.ReadLine());   // use float, not int
                Console.Write("Input the Department: ");
                dept = Console.ReadLine();
                Console.Write("Input the Description: ");
                description = Console.ReadLine();

                Console.WriteLine("Receipt Info:");
                Console.WriteLine($"Date: {date}\nTotal: {total}\nDepartment: {dept}\nDescription: {description}");
                Console.Write("Is this Correct? (Y/N): ");

                if (Console.ReadLine().ToUpper().Equals("Y"))
                {
                    control.AddReceipt(dept, date, description, total);
                    Console.WriteLine("Receipt staged for adding...");
                }
                else
                {
                    Console.WriteLine("Canceled...");
                }

                Console.Write("Add another receipt? (Y/N): ");
                if (Console.ReadLine().ToUpper().Equals("N"))
                {
                    break;
                }
            }

            // ✅ batch push like in AddReceipt()
            Console.WriteLine("Press enter to push receipts to Excel...");
            string x = Console.ReadLine();
            if (x == "")
            {
                Console.WriteLine("Pushing to Excel...");

                foreach (var rec in control.receipts)
                {
                    control.WriteToExel(rec);
                    Console.WriteLine("Added: " + rec);
                }

                control.receipts.Clear();
            }

            Console.WriteLine("\n\n\n\n\n\n\n\n");
        }

        public static void sort()
        {
            while (true) 
            { 
                Console.WriteLine("Sorting Methods: \nSort by Date: 1)\nSort by Total: 2)\nSort by Dept: 3)");
                string x = Console.ReadLine();
                if (x.Equals("1"))
                {
                    control.SortExelByDate();
                    Console.WriteLine("Done Sorting");
                    control.receipts.Clear();
                    break;
                }
                else if (x.Equals("2"))
                {
                    control.SortExelByTotal();
                    Console.WriteLine("Done Sorting");
                    control.receipts.Clear();
                    break;
                }
                else if (x.Equals("3"))
                {
                    // Define department options
                    string[] deptOptions = {
                        "Treasure",   // 1
                        "Social",     // 2
                        "Philo",      // 3
                        "Recruitment",// 4
                        "Historian",  // 5
                        "Chaplan",    // 6
                        "House"       // 7
                };

                    // Display menu
                    Console.WriteLine("Which department should come first?");
                    for (int i = 0; i < deptOptions.Length; i++)
                    {
                        Console.WriteLine($"{i + 1}) {deptOptions[i]}");
                    }

                    Console.Write("Enter number (1-7): ");
                    string input = Console.ReadLine();

                    // Validate input
                    if (int.TryParse(input, out int choice) && choice >= 1 && choice <= deptOptions.Length)
                    {
                        string selectedDept = deptOptions[choice - 1];
                        control.SortExelByDepartmentFirst(selectedDept);
                    }
                    else
                    {
                        Console.WriteLine("Invalid selection. No sorting applied.");
                    }
                    control.receipts.Clear();

                }
                break;
            }
        }

    }
}
