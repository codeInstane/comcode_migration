using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Data;
using System.Data.OleDb;
using System.IO;

namespace comcode_migration
{
    class cmdline_migrate
    {    

        public static void cmdline_migration(string filePath)        {

            // Check if file path and name provided as command line argument
            //if (args.Length == 0)
            //{
            //    Console.WriteLine("Please provide the file path and name as a command line argument.");
            //    return;
            //}

            //string filePath = args[0];

            // Connection string for Excel file
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;'";

            // Create connection and command objects
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand("SELECT * FROM [Sheet1$]", connection);

                // Open connection and create data reader
                connection.Open();
                OleDbDataReader reader = command.ExecuteReader();

                // Create output file and writer
                using (StreamWriter writer = new StreamWriter("output.txt"))
                {
                    // Loop through rows and write to file
                    while (reader.Read())
                    {
                        // Write each column value to file separated by tabs
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            writer.Write(reader.GetValue(i) + "\t");
                                Console.WriteLine();
                            }
                        writer.WriteLine();
                    }
                }

                // Close data reader and connection
                reader.Close();
            }
        }
    }













}

