using System;
using System.Data.OleDb;
using System.IO;

/* References:
 * 
 * OleDbDataReader Class: https://docs.microsoft.com/en-us/dotnet/api/system.data.oledb.oledbdatareader?view=netframework-4.7.2 
 * 
 * Copy, Delete, and move files and folders: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/file-system/how-to-copy-delete-and-move-files-and-folders
 * 
 * Query for Files with a specified Attrribute or Name: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/concepts/linq/how-to-query-for-files-with-a-specified-attribute-or-name
 * 
 * Get Information About Files, Folders, and Drives: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/file-system/how-to-get-information-about-files-folders-and-drives
 * 
 * Get files/folder names with a partial file names: https://social.msdn.microsoft.com/Forums/windows/en-US/f95a51fa-d1ac-44f4-874e-6d1b6c46282f/how-to-find-folders-and-files-by-its-partial-name-c?forum=csharpgeneral
 
 */


namespace FileMover
{

    /* Note: This application moves files with the following conditions: 
     * 1. Record in Access Database and the corresponding pdf file should exist in a folder 
     * 2. File has a correct format ("Registration#_DocumentID") 
     * 3. Record has an Expiration Date 
     */
    class Program
    {
        static OleDbConnection connection;
        static string registrationNumber;
        static FileInfo[] filesInDirs;

        static string sourcePath = @"C:\Users\skang\Documents\TradeMarks\Access_Images";
        static string connectionString = @"Provider =Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\skang\Documents\TMTM100\TMTM10_V1.3.mdb";

        //Strings for query clauses 
        static string queryTmActive = "Select RegistrationType, RegistrationNumber, ExpirationDate from Registration where RegistrationType='TM' and ExpirationDate is not null and ExpirationDate >= now() ";
        static string queryTmInactive = "Select RegistrationType, RegistrationNumber, ExpirationDate from Registration where RegistrationType='TM' and ExpirationDate is not null and ExpirationDate < now()";
        static string querySmActive = "Select RegistrationType, RegistrationNumber, ExpirationDate from Registration where RegistrationType='SM' and ExpirationDate is not null and ExpirationDate >= now() ";
        static string querySmInactive = "Select RegistrationType, RegistrationNumber, ExpirationDate from Registration where RegistrationType='SM' and ExpirationDate is not null and ExpirationDate < now() ";
        static string queryTmNullExp = "Select RegistrationType, RegistrationNumber, ExpirationDate from Registration where RegistrationType='TM' and ExpirationDate is null";

        //static string queryRenewal = "Select * from DocHistory where RegistrationType='TM' and TypeOfDoc ='RENWL' and FilingDate > DateAdd(year, -5, now())";

        //String folder paths for TradeMarks & ServiceMarks
        static string activeTmPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\TMs\Active_TMs";
        static string inActiveTmPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\TMs\InActive_TMs";
        static string activeSmPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\SMs\Active_SMs";
        static string inActiveSmPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\SMs\InActive_SMs";


        static string tmQuery = "Select * from Registration where RegistrationType='TM'";
        static string smQuery = "Select * from Registration where RegistrationType='SM'";
        static string tmPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\TMs";
        static string smPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\SMs";
        static string tmEmptyPath = @"C:\Users\skang\Documents\TradeMarks\Access_Images\TMs_Empty_Exp";


        static void Main(string[] args)
        {

            //Create a connection
            using (connection = new OleDbConnection(connectionString))
            {

                //MoveFiles(tmQuery, tmPath);
                //MoveFiles(smQuery, smPath);
                //MoveFiles(queryTmActive, activeTmPath);
                //MoveFiles(queryTmInactive, inActiveTmPath);
                //MoveFiles(querySmActive, activeSmPath);
                //MoveFiles(querySmInactive, inActiveSmPath);
                //MoveFiles(queryTmNullExp, tmEmptyPath);
                CountFiles(queryTmActive);

                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }

        }

        //Move PDF files to a destination folder 
        static void MoveFiles(string queryString, string folderPath)
        {
            // Create a command and set its connection  
            OleDbCommand command = new OleDbCommand(queryString, connection);
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"C:\Users\skang\Documents\TradeMarks\Access_Images");
            // Open the connection and execute the select command.  
            try
            {
                connection.Open();
                // Execute command and retrieve records  
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    var count = 0;
                    //Console.WriteLine("------------Registration Records----------------");
                    while (reader.Read())
                    {
                        //Console.WriteLine("{0} {1} {2}", reader["RegistrationType"].ToString(), reader["RegistrationNumber"].ToString(), reader["ExpirationDate"].ToString());
                        //Console.WriteLine("{0} {1} {2}", reader[0].ToString(), reader[1].ToString() ,reader[18].ToString());

                        registrationNumber = reader["RegistrationNumber"].ToString();
                        filesInDirs = hdDirectoryInWhichToSearch.GetFiles("*" + registrationNumber + "*_*");


                        foreach (FileInfo foundFile in filesInDirs)
                        {
                            string fullName = foundFile.FullName;
                            string fileName = Path.GetFileName(fullName);


                            Console.WriteLine(registrationNumber);


                            //Use Path class to manipulate file and directory paths.
                            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                            string destFile = System.IO.Path.Combine(folderPath, fileName);

                            System.IO.File.Move(sourceFile, destFile);
                            count++;
                        }

                    }
                    Console.WriteLine(count);
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /* Counting records 
         * 
         */
        static void CountFiles(string queryString)
        {
            // Create a command and set its connection  
            OleDbCommand command = new OleDbCommand(queryString, connection);
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"C:\Users\skang\Documents\TradeMarks\Access_Images");
            // Open the connection and execute the select command.  
            try
            {
                connection.Open();
                // Execute command and retrieve records  
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    var count = 0;
                    //Console.WriteLine("------------Registration Records----------------");
                    while (reader.Read())
                    {

                        Console.WriteLine("{0}", reader[1].ToString());

                        count++;
                    }
                    Console.WriteLine(count);
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    } //Class Program
}
