using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Exceltodb2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ExceltoDbTable();
        }
        // You will enter the path which your excel file exist. 
        public static string path = @"C:\Users\styunust\Desktop\2020 Spring Courses\PURE\PureData\TürkiyeDiziDatası\1956Imdb4.2.xlsx";
        //There is a two type of connectrion string.First of them is for xlsx files,second of them is for xls files
        static string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes'", path);
        public static string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        public static DataSet ds;
        public static void ExceltoDbTable()
        {
            OleDbConnection OleDbcon = new OleDbConnection(connString); //We are connecting the excel file
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]",OleDbcon); //Creating the command.It will select the all of the file
//           OleDbDataAdapter objAdapter1 = new OleDbDataAdapter(cmd); //objAdapter datayı selectliyor üstteki komut ile
//           ds = new DataSet(); 
//           objAdapter1.Fill(ds);
            OleDbcon.Open(); // Opening our excel file
            DbDataReader dr = cmd.ExecuteReader(); //We select the data and take the rows from datasource
            string constr = @"Server=(localdb)\ProjectsV13;Database=TvSeriesDb;Trusted_Connection=true"; 
            SqlBulkCopy bulkInsert = new SqlBulkCopy(constr); // taking referans of the db
            bulkInsert.DestinationTableName = "TvSeries2"; //finding the table which will run the command
            bulkInsert.WriteToServer(dr); // inserting our selected excel to our table
            OleDbcon.Close();





        }
    }
}
