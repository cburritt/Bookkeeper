using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Text.RegularExpressions;

using Tesseract;

public class Control()
{
    public List<Receipt> receipts = new List<Receipt>();


    string filePath = @"C:\Users\cburr\Desktop\exelTester\Students.xlsx";

    public void AddReceipt(string dept, string date, string description, float totalCost)
    {
        Receipt receipt = new Receipt(dept, date, description, totalCost);
        receipts.Add(receipt);
    }
    public void RemoveReceipt(Receipt receipt)
    {
        receipts.Remove(receipt);
    }
    public List<Receipt> SearchAllReceipts()
    {
        List<Receipt> result = new List<Receipt>();
        if (receipts.Count == 0)
        {
            Console.WriteLine("No receipts to print.");
            return null;
        }
        

        // Sort receipts by TotalCost in ascending order
        

        foreach (var receipt in receipts)
        {
            result.Add(receipt);
        }
        result.Sort((r1, r2) => r1.totalCost.CompareTo(r2.totalCost));
        return result;

    }
    public List<Receipt> SearchReceiptsByDate(string date)
    {
        List<Receipt> result = new List<Receipt>();

        
        string[] parts = date.Split('/');
        if (parts.Length != 3)
        {
            throw new ArgumentException("Date must be in the format MM/DD/YY or MM/DD/YYYY");
        }

        int month = int.Parse(parts[0]);
        int day = int.Parse(parts[1]);
        int year = int.Parse(parts[2]);
        if (year < 100)
        {
            year += 2000;
        }

        
        foreach (var receipt in receipts)
        {
            if (receipt.month == month && receipt.day == day && receipt.year == year)
            {
                result.Add(receipt);
            }
        }

        return result;
    }
    public List<Receipt> SearchReceiptsByMonth(int month)
    {
        List<Receipt> result = new List<Receipt>();
        
        foreach (var receipt in receipts)
        {
            if (receipt.month == month)
            {
                result.Add(receipt);
            }
        }

        return result;
    }
    public List<Receipt> SearchReceiptsByDay(int year)
    {
        List<Receipt> result = new List<Receipt>();
        
        foreach (var receipt in receipts)
        {
            if (receipt.year == year)
            {
                result.Add(receipt);
            }
        }

        return result;
    }

    public List<Receipt> SearchReceiptsByDept(string dept)
    {
        List<Receipt> result = new List<Receipt>();

        foreach (var receipt in receipts)
        {
            if (receipt.dept == dept)
            {
                result.Add(receipt);
            }
        }

        return result;
    }

    /// <summary>
    /// Writes a Recipt to exel
    /// </summary>
    /// <param name="receipt"></param>
    public  void WriteToExel(Receipt receipt)
    {
        

        // Open if exists, else create new
        XLWorkbook workbook;
        if (System.IO.File.Exists(filePath))
        {
            workbook = new XLWorkbook(filePath);
        }
        else
        {
            workbook = new XLWorkbook();
        }

        // Get or create worksheet
        var worksheet = workbook.Worksheet("Students");
        if (worksheet == null)
        {
            worksheet = workbook.Worksheets.Add("Students");
            // Add headers if sheet is new
            worksheet.Cell(1, 1).Value = "Date";
            worksheet.Cell(1, 2).Value = "Dept";
            worksheet.Cell(1, 3).Value = "Total";
            worksheet.Cell(1, 4).Value = "Description";
        }

        // Find the next empty row
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        int newRow = lastRow + 1;

        // Example data you want to write
        worksheet.Cell(newRow, 1).Value = receipt.date;
        worksheet.Cell(newRow, 2).Value = receipt.dept;
        worksheet.Cell(newRow, 3).Value = receipt.totalCost;
        worksheet.Cell(newRow, 4).Value = receipt.description;

        // Save back to file
        workbook.SaveAs(filePath);

        Console.WriteLine("Excel file updated: " + filePath);
    }

    public void CollectExel()
    {
        if (!System.IO.File.Exists(filePath))
        {
            Console.WriteLine("Excel file does not exist.");
            return;
        }

        var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet("Students");

        if (worksheet == null)
        {
            Console.WriteLine("Worksheet 'Students' does not exist.");
            return;
        }

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

        receipts.Clear(); // reset before collecting

        for (int i = 2; i <= lastRow; i++) // skip headers
        {
            string date = worksheet.Cell(i, 1).GetValue<string>();
            string dept = worksheet.Cell(i, 2).GetValue<string>();
            float totalCost = worksheet.Cell(i, 3).GetValue<float>();
            string description = worksheet.Cell(i, 4).GetValue<string>();

            AddReceipt(dept, date, description, totalCost);
        }

        Console.WriteLine("Receipts collected from Excel.");
    }


    public string ReadReceiptDate(string imagePath)
    {
        string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tess");

        using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
        using (var img = Pix.LoadFromFile(imagePath))
        using (var page = engine.Process(img))
        {
            string text = page.GetText();

            // Regex to capture common date-like patterns
            var dateMatch = Regex.Match(text, @"\b(\d{1,4}[/-]\d{1,2}[/-]\d{1,4})\b");

            if (dateMatch.Success)
            {
                string rawDate = dateMatch.Groups[1].Value;

                // Accept multiple possible date formats
                string[] formats = {
                "MM/dd/yy", "MM/dd/yyyy",
                "M/d/yy",  "M/d/yyyy",
                "MM-dd-yy", "MM-dd-yyyy"
                
               
           
            };

                if (DateTime.TryParseExact(rawDate, formats,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime parsedDate))
                {
                    // Normalize format into what SetDate expects
                    return parsedDate.ToString("MM/dd/yyyy");
                }

                // If parsing fails, return raw OCR text
                return rawDate;
            }
            else
            {
                // Fallback date if OCR fails
                return "12/12/2005";
            }
        }
    }

    


    public string ReadReceiptTotal(string imagePath)
    {
        string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tess");

        using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
        using (var img = Pix.LoadFromFile(imagePath))
        using (var page = engine.Process(img))
        {
            string text = page.GetText();
            

            
            var totalMatch = Regex.Match(
                text,
                @"\btotal\b(?!\s*(sub|tax|discount))[:\s]*\$?\s*([\d.,]+)",
                RegexOptions.IgnoreCase
            );

            if (totalMatch.Success)
            {
                return totalMatch.Groups[2].Value; // <-- the actual number
            }
            else
            {
                return "No total found";
            }
        }
    }




    public void SortExelByDate()
    {
        // Load all receipts from Excel
        CollectExel();

        // Sort by date
        receipts = receipts
            .OrderBy(r => DateTime.ParseExact(r.date, "MM/dd/yyyy", null))
            .ToList();

        // Rewrite Excel from scratch
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Students");

        // headers
        worksheet.Cell(1, 1).Value = "Date";
        worksheet.Cell(1, 2).Value = "Dept";
        worksheet.Cell(1, 3).Value = "Total";
        worksheet.Cell(1, 4).Value = "Description";

        int row = 2;
        foreach (var r in receipts)
        {
            worksheet.Cell(row, 1).Value = r.date;
            worksheet.Cell(row, 2).Value = r.dept;
            worksheet.Cell(row, 3).Value = r.totalCost;
            worksheet.Cell(row, 4).Value = r.description;
            row++;
        }

        workbook.SaveAs(filePath);
        
    }

    public void SortExelByTotal()
    {
        receipts = receipts.OrderByDescending(r => r.totalCost).ToList();
        foreach (var rec in receipts)
        {
            WriteToExel(rec);
            
        }
    }


    public void SortExelByDepartmentFirst(string primaryDept)
    {
        string[] deptOrder = {
        "Treasure", "Social", "Philo", "Recruitment", "Historian", "Chaplan", "House"
    };

        // Put the chosen department first
        receipts = receipts
            .OrderBy(r => r.dept.Equals(primaryDept, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ThenBy(r => Array.IndexOf(deptOrder, r.dept))
            .ToList();

        // Write sorted receipts to Excel
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Students");

        worksheet.Cell(1, 1).Value = "Date";
        worksheet.Cell(1, 2).Value = "Dept";
        worksheet.Cell(1, 3).Value = "Total";
        worksheet.Cell(1, 4).Value = "Description";

        int row = 2;
        foreach (var r in receipts)
        {
            worksheet.Cell(row, 1).Value = r.date;
            worksheet.Cell(row, 2).Value = r.dept;
            worksheet.Cell(row, 3).Value = r.totalCost;
            worksheet.Cell(row, 4).Value = r.description;
            row++;
        }

        workbook.SaveAs(filePath);
        Console.WriteLine($"Excel sorted with {primaryDept} first and saved: {filePath}");
    }

}
