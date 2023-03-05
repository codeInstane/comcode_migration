using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using ExcelDataReader;
using System.IO;
using static comcode_migration.comcode;
using System.Diagnostics;
using static comcode_migration.comcodetable;

namespace comcode_migration
{
    class comcode
    {

        // Connection string for SQL Server
        //private readonly String ConnStr = @"Data Source=tcp:sqlsrv-4s-sit-001.database.windows.net,1433;Initial Catalog = sqldb-4s-sit;Persist Security Info=False;User ID = AzureServerAdmin; Password=Password123;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;";

        // Path to the Excel file
        // private readonly String filePath = "C:\\myExcelFile.xlsx";



        internal class InsertDataFromExcel
        {
            public List<comcodetable> ReadFromExcelfile(string filePath)
            {

                List<comcodetable> person = new List<comcodetable>();

                //string filePath = "C:\\Users\\staff\\Documents\\computer programming\\C#\\comcode_migration\\myExcelFile.xlsx";

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {                 
                    // Read data from Excel file
                    
                    using (var reader = ExcelReaderFactory.CreateReader(stream))

                    {
                        // Get the first worksheet in the Excel file
                        reader.Read();
                        
                        DataTable dataTable = reader.AsDataSet().Tables[0];

                        for (int i = 1; i < dataTable.Rows.Count; i++)
                        {

                            comcodetable personal = new comcodetable

                            {
                                FirstName = dataTable.Rows[i][0].ToString(),
                                //Console.WriteLine(FirstName);

                                LastName = dataTable.Rows[i][1].ToString(),
                                //Console.WriteLine(lastname);

                                Category = dataTable.Rows[i][2].ToString(),

                                //Console.WriteLine(category);
                            };

                            person.Add(personal);
                            //Console.WriteLine(personal);
                        }
                    }
                    return person;
                }
            }

            internal void InsertDataToDB(List<comcodetable> people)
            {

                try
                {


                    string connectString = @"Data Source=DESKTOP-EVDH83E\SQLEXPRESS;
                                    Initial Catalog = persons;
                                    Trusted_Connection=True";
                    //  Persist Security Info=False;
                    //   User ID = AzureServerAdmin; Password=Password123;
                    //   MultipleActiveResultSets=False;Encrypt=True;
                    //   TrustServerCertificate=False";




                    SqlConnectionStringBuilder builder =
                        new SqlConnectionStringBuilder(connectString);
                    Console.WriteLine("Original: " + builder.ConnectionString);

                    //  DataSource = "HOIT2105NB08\\SQLEXPRESS",
                    //     InitialCatalog = "person",

                    //UserID = "AzureServerAdmin",
                    //     Password = "Password123",


                    //
                    //}.ToString()))


                    using (SqlConnection connection =
                           new SqlConnection(builder.ConnectionString))
                    {
                        connection.Open();
                        // Now use the open connection.
                        Console.WriteLine("Database = " + connection.Database);

                        int migr = 0;
                        int dups = 0;

                        foreach (var person in people)
                        {
                            //Console.WriteLine("enter foreach loop");
                            string firstname = person.FirstName;
                            string lastname = person.LastName;
                            string category = person.Category;

                           // Console.WriteLine(firstname); it works

                            var command = new SqlCommand("SELECT COUNT(*) FROM persons WHERE FirstName = @FirstName AND Category = @Category", connection);
                            command.Parameters.AddWithValue("@FirstName", firstname);
                            command.Parameters.AddWithValue("@Category", category);

                           // Console.WriteLine(firstname); it works

                            var count = (int)command.ExecuteScalar();
                            if (count == 0)
                            {
                                //Console.WriteLine("enter if count ==0 ");    
                                command = new SqlCommand("INSERT INTO persons (FirstName, LastName, Category) VALUES (@FirstName, @LastName, @Category)", connection);
                                command.Parameters.AddWithValue("@FirstName", firstname);
                                command.Parameters.AddWithValue("@LastName", lastname);
                                command.Parameters.AddWithValue("@Category", category);
                                command.ExecuteNonQuery();

                                migr++;
                                Console.WriteLine($"Now mirgrating: \t{firstname}\t{lastname}\t{category}");                                
                                //Debug.WriteLine("Now mirgrating", firstname + "  " + lastname + "   " + category);   
                            }
                            else
                            {                                
                                Console.WriteLine($"Duplicates comcode: \t{firstname}\t{lastname}\t{category}");
                                dups++;                                
                                //Debug.WriteLine("Duplicates comcode", firstname + "  " + lastname + "   " + category);
                            }
                        }
                        Console.WriteLine($"Total number of records mirgate :\t{ migr}");
                        Console.WriteLine($"Total number of records duplicate :\t{ dups}");

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                
                //throw new NotImplementedException();
            }
        }

                 

        public static void Main(string[] args)

        {

            // Check if file path and name provided as command line argument
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the file path and name as a command line argument.");
                return;
            }

            string filePath = args[0];

            InsertDataFromExcel Idf = new InsertDataFromExcel();

            var personals = Idf.ReadFromExcelfile(filePath);

            //    Console.WriteLine("Finishing read comcode excel template file");

            //    //foreach (var peoples in personals)
            //    //{
            //    //    Console.WriteLine(peoples.FirstName + "   " + peoples.LastName + "   " + peoples.Category);
            //    //}
            //    //Console.WriteLine(personals.Count);

                Idf.InsertDataToDB(personals);


            //    Console.WriteLine("Finishing Comcode migration.");
            //    Console.ReadLine();



            //cmdline_migrate.cmdline_migration(filePath);


            //Itest_connect.Migratedb();            

        }
    }
    }




    

    

