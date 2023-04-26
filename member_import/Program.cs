using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string excelFilePath = "path/to/your/excel/file.xlsx";
        string connectionString = "Data Source=(local);Initial Catalog=YourDatabaseName;Integrated Security=True"; // Replace with your actual connection string

        // Read data from Excel file using EPPlus
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first sheet of the Excel file
            int rows = worksheet.Dimension.Rows;

            // Store data into SQL Server database
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int row = 2; row <= rows; row++) // Start from row 2 to skip header row
                {
                    // Generate a unique 8-digit MemberID
                    string memberId = GenerateMemberId();

                    // Extract data from the row
                    string firstname = worksheet.Cells[row, 2].Value?.ToString();
                    string lastname = worksheet.Cells[row, 3].Value?.ToString();
                    string middlename = worksheet.Cells[row, 4].Value?.ToString();
                    DateTime dob = Convert.ToDateTime(worksheet.Cells[row, 5].Value);
                    string gender = worksheet.Cells[row, 6].Value?.ToString();
                    string addr1 = worksheet.Cells[row, 7].Value?.ToString();
                    string addr2 = worksheet.Cells[row, 8].Value?.ToString();
                    string city = worksheet.Cells[row, 9].Value?.ToString();
                    string zip = worksheet.Cells[row, 10].Value?.ToString();
                    string ssn = worksheet.Cells[row, 11].Value?.ToString();
                    string createId = "YourSQLUsername"; // Replace with your actual SQL username
                    DateTime createDate = DateTime.Now;
                    string updateId = "YourSQLUsername"; // Replace with your actual SQL username
                    DateTime lastUpdate = DateTime.Now;

                    // Insert data into SQL Server
                    string query = @"INSERT INTO YourTableName (MemberId, Firstname, Lastname, MiddleName, DOB, Gender, Addr1, Addr2, City, Zip, SSN, CreateID, CreateDate, UpdateID, LastUpdate)
                                     VALUES (@MemberId, @Firstname, @Lastname, @MiddleName, @DOB, @Gender, @Addr1, @Addr2, @City, @Zip, @SSN, @CreateID, @CreateDate, @UpdateID, @LastUpdate)";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@MemberId", memberId);
                        command.Parameters.AddWithValue("@Firstname", firstname);
                        command.Parameters.AddWithValue("@Lastname", lastname);
                        command.Parameters.AddWithValue("@MiddleName", middlename);
                        command.Parameters.AddWithValue("@DOB", dob);
                        command.Parameters.AddWithValue("@Gender", gender);
                        command.Parameters.AddWithValue("@Addr1", addr1);
                        command.Parameters.AddWithValue("@Addr2", addr2);
                        command.Parameters.AddWithValue("@City", city);
                        command.Parameters.AddWithValue("@Zip", zip);
                        command.Parameters.AddWithValue("@SSN", ssn);
                        command.Parameters.AddWithValue("@CreateID", createId);
                        command.Parameters.AddWithValue("@CreateDate", createDate);
                        command.Parameters.AddWithValue("@UpdateID", updateId);
                        command.Parameters.AddWithValue("@LastUpdate", lastUpdate);

                        command.ExecuteNonQuery();
                    }
                }

                Console.WriteLine("Data has been inserted into SQL Server successfully.");
            }

            }
        }

    private static string GenerateMemberId() { 

        string memberId = Guid.NewGuid().ToString().Substring(0, 8);
        return memberId;
        }

    }





