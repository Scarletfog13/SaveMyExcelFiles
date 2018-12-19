using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//added below name spaces
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace SaveExcel
{
    
    class Program
    {
        static String InputTableName;
        static String Path;
        static String DbName;
        static String ErrorFolder;
        static void Main(string[] args)
        {
            Console.WriteLine("Enter File Path ");
            Path = Console.ReadLine();
            Console.WriteLine("Enter Table name: ");
            InputTableName = Console.ReadLine();          
            Console.WriteLine("Enter Database name ");
            DbName = Console.ReadLine();
            Console.WriteLine("Specify Folder to save error in case ");
            ErrorFolder = Console.ReadLine();
            //the datetime and Log folder will be used for error log file in case error occured
            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            string LogFolder = ErrorFolder;
            try
            {
                //Provide the Source Folder path where excel files are present
                String FolderPath = @Path;
                //Provide the Database Name 
                string DatabaseName = DbName;
                //Provide the SQL Server Name 
                string SQLServerName = "(local)";
                //Provide the table name in which you want to load excel sheet's data
                 string TableName = @InputTableName;
                //Provide the schema of table
                String SchemaName = @"dbo";


                //Create Connection to SQL Server Database 
                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = "Data Source = "
                    + SQLServerName + "; Initial Catalog ="
                    + DatabaseName + "; "
                    + "Integrated Security=true;";

                var directory = new DirectoryInfo(FolderPath);
                FileInfo[] files = directory.GetFiles();

                //Declare and initilize variables
                string fileFullPath = "";


                //Get one Book(Excel file at a time)
                foreach (FileInfo file in files)
                {
                    fileFullPath = FolderPath + "\\" + file.Name;

                    //Create Excel Connection
                    string ConStr;
                    string HDR;
                    HDR = "YES";
                    ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                        + fileFullPath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                    OleDbConnection cnn = new OleDbConnection(ConStr);

                    //Get Sheet Name
                    cnn.Open();
                    DataTable dtSheet = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetname;
                    sheetname = "";

                    //Loop through each sheet
                    foreach (DataRow drSheet in dtSheet.Rows)
                    {
                        if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                        {
                            sheetname = drSheet["TABLE_NAME"].ToString();

                            //Get data from Excel Sheet to DataTable
                            OleDbConnection Conn = new OleDbConnection(ConStr);
                            Conn.Open();
                            OleDbCommand oconn = new OleDbCommand("select * from [" + sheetname + "]", Conn);
                            OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                            DataTable dt = new DataTable();
                            adp.Fill(dt);
                            Conn.Close();

                            SQLConnection.Open();
                            //Load Data from DataTable to SQL Server Table.
                            using (SqlBulkCopy BC = new SqlBulkCopy(SQLConnection))
                            {
                                BC.BulkCopyTimeout = 3600;
                                BC.DestinationTableName = SchemaName + "." + TableName;
                                foreach (var column in dt.Columns)
                                    BC.ColumnMappings.Add(column.ToString(), column.ToString());

                               
                                BC.WriteToServer(dt);
                             
                            }
                            SQLConnection.Close();
                        }
                       
                    }
                    
                }
                
            }
            catch (Exception exception)
            {
                // Create Log File for Errors
                using (StreamWriter sw = File.CreateText(LogFolder
                    + "\\" + "ErrorLog_" + datetime + ".log"))
                {
                    sw.WriteLine(exception.ToString());

                }

            }

        }
    }
}