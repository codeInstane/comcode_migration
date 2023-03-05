using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using ExcelDataReader;
using System.IO;

namespace comcode_migration
{
    class Itest_connect
    {

        // Connection string for SQL Server
        //private readonly String connectionString = @"Data Source=tcp:sqlsrv-4s-sit-001.database.windows.net,1433;Initial Catalog = sqldb-4s-sit;Persist Security Info=False;User ID = AzureServerAdmin; Password=Password123;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;";

        // Path to the Excel file
        //private readonly String filePath = "C:\\myExcelFile.xlsx";

        
            public static void Migratedb()
            {
            String filePath = "C:\\Users\\staff\\Documents\\computer programming\\C#\\comcode_migration\\myExcelFile.xlsx";
            //String connectionString = @"Data Source=tcp:sqlsrv-4s-sit-001.database.windows.net,1433;
            //                            Initial Catalog = sqldb-4s-sit;
            //                            Persist Security Info=False;
            //                            User ID = AzureServerAdmin; 
            //                            Password=Password123;
            //                            MultipleActiveResultSets=False;
            //                            Encrypt=True;TrustServerCertificate=False;";

            string connectStr = @"Data Source=DESKTOP-EVDH83E\SQLEXPRESS;
                                    Initial Catalog = persons;
                                    Trusted_Connection=True";

            // Open the Excel file and read the data
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                    // Get the first worksheet in the Excel file
                    reader.Read();
                    //var result = reader.AsDataSet().Tables[0];
                       DataTable dataTable = reader.AsDataSet().Tables[0];
                    //Console.WriteLine(dataTable.Rows[1][0]);
                        // Connect to the SQL Server database
                        using (SqlConnection connection = new SqlConnection(connectStr))
                        {
                            connection.Open();

                        // Iterate through the rows in the Excel file and insert them into the database
                        //foreach (DataRow row in dataTable.Rows)

                        for (int i = 1; i < dataTable.Rows.Count; i++)
                        {
                            string firstname = dataTable.Rows[i][0].ToString();
                            Console.WriteLine(firstname);

                            string lastname = dataTable.Rows[i][1].ToString();
                            Console.WriteLine(lastname);

                            string category = dataTable.Rows[i][2].ToString();
                            Console.WriteLine(category);

                            // Get the firstname value from the Excel row
                            //string firstname = row["FirstName"].ToString();


                            //string firstname = "FirstName" ;
                            // Check if the firstname value already exists in the database
                            using (SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM persons WHERE FirstName=@FirstName", connection))
                                {
                                    checkCommand.Parameters.AddWithValue("@FirstName", firstname);
                                    
                                    int count = (int)checkCommand.ExecuteScalar();

                                    // If the firstname value does not exist in the database, insert the row
                                    if (count == 0)
                                    {
                                        using (SqlCommand insertCommand = new SqlCommand("INSERT INTO persons (FirstName) VALUES (@FirstName)", connection))
                                        {
                                            insertCommand.Parameters.AddWithValue("@FirstName", firstname);
                                            insertCommand.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            Console.WriteLine("Press any key to finish.");

            Console.ReadLine();

        }
        }
    }
