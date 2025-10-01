/*
 * Author: Caden Burritt
 */
public class Receipt
{
    public string dept;
    public string date;
    public int month;
    public int day;
    public int year;
    public string description;
    public float totalCost;

    /// <summary>
    /// Constructs Receipt
    /// </summary>
    /// <param name="dept"></param>
    /// <param name="date"></param>
    /// <param name="description"></param>
    /// <param name="totalCost"></param>
    public Receipt(string dept, string date, string description, float totalCost)
    {
        this.dept = dept;
        this.date = date;
        this.description = description;
        this.totalCost = totalCost;
        SetDate(date);
    }

    /// <summary>
    /// Constructor for empty input
    /// </summary>
    public Receipt()
    {
        this.dept = "no input";
        this.date = "no date";
        this.description = "no description";
        this.totalCost = 0;
        this.month = 0;
        this.day = 0;
        this.year = 0;
    }

    /// <summary>
    /// Parses a date string in MM/DD/YYYY or MM/DD/YY format and sets the month, day, and year.
    /// </summary>
    private void SetDate(string date)
    {
        // Extract only the date portion (if OCR includes time)
        date = date.Split(' ')[0];

        string[] parts = date.Split('/');
        if (parts.Length != 3)
        {
            throw new ArgumentException("Date must be in the format MM/DD/YY or MM/DD/YYYY");
        }

        this.month = int.Parse(parts[0]);
        this.day = int.Parse(parts[1]);
        this.year = int.Parse(parts[2]);

        if (this.year < 100)
        {
            this.year += 2000;
        }
    }

    /// <summary>
    /// Converts the object to a string representation.
    /// </summary>
    public override string ToString()
    {
        return $"Date: {this.month}/{this.day}/{this.year}, Dept: {this.dept}, Description: {this.description}, Total Cost: ${this.totalCost}";
    }
}