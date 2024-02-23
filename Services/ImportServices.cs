using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

namespace ExcelFileHandling_WebTech.Services
{
    public class ImportServices
    {
        public void ProcessExcelData(string tableName, DataTable data, string worksheetName)
        {
            string fullTableName = $"{tableName}_{worksheetName}";
            CreateTableInDatabase(fullTableName, data.Columns);
            InsertDataIntoTable(fullTableName, data);
        }

        private void CreateTableInDatabase(string tableName, DataColumnCollection columns)
        {
            // Utilize the connection string from Web.Config
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["ExcelDBConnectionString"].ConnectionString;

            // Create a connection
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Create a command to check if the table already exists
                using (SqlCommand command = new SqlCommand($"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}') BEGIN CREATE TABLE {tableName} ({GetTableColumns(columns)}) END", connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private void InsertDataIntoTable(string tableName, DataTable data)
        {
            // Utilize the connection string from Web.Config
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["ExcelDBConnectionString"].ConnectionString;

            // Create a connection
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Loop through rows and insert data into the table
                foreach (DataRow row in data.Rows)
                {
                    // Create a command for each row
                    using (SqlCommand command = new SqlCommand($"INSERT INTO {tableName} VALUES ({GetRowValues(row)})"))
                    {
                        command.Connection = connection;
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public DataTable ExtractDataFromExcel(ExcelWorksheet worksheet)
        {
            DataTable table = new DataTable();
            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }

            for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var worksheetRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                DataRow row = table.Rows.Add();
                foreach (var cell in worksheetRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }

            return table;
        }

        private string GetTableColumns(DataColumnCollection columns)
        {
            // Determine column data types based on the data in the first row
            List<string> columnDefinitions = new List<string>();
            foreach (DataColumn column in columns)
            {
                string columnName = column.ColumnName;
                Type dataType = column.DataType;

                // Skip RowNumber column
                if (columnName.Equals("RowNumber", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Choose SQL data type based on .NET data type
                string sqlType;
                if (dataType == typeof(string))
                {
                    sqlType = "NVARCHAR(MAX)";
                }
                else if (dataType == typeof(int))
                {
                    sqlType = "INT";
                }
                else if (dataType == typeof(DateTime))
                {
                    sqlType = "DATETIME";
                }
                else
                {
                    // Adjust or add more cases as needed
                    sqlType = "NVARCHAR(MAX)";
                }

                // Add column definition to the list
                columnDefinitions.Add($"{columnName} {sqlType}");
            }

            return string.Join(", ", columnDefinitions);
        }

        private string GetRowValues(DataRow row)
        {
            // Convert row values to SQL-friendly string
            List<string> values = new List<string>();
            foreach (var value in row.ItemArray)
            {
                // Escape single quotes and convert to string
                values.Add($"'{value?.ToString()?.Replace("'", "''")}'");
            }

            return string.Join(", ", values);
        }
    }
}