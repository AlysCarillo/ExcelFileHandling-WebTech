# Excel File Handling in ASP.NET MVC

This ASP.NET MVC project demonstrates how to upload, extract, and process data from Excel files, creating tables in a SQL Server database for each worksheet in the Excel file.

## Getting Started

### Prerequisites

- Visual Studio (or any preferred IDE for ASP.NET MVC development)
- SQL Server (Express or higher)

### Setup

1. Clone the repository to your local machine.
2. Open the project in Visual Studio.
3. Update the connection string in the `Web.config` file to point to your SQL Server instance.

    ```xml
    <connectionStrings>
        <add name="ExcelDBConnectionString" connectionString="Data Source=YourServer;Initial Catalog=YourDatabase;Integrated Security=True;" providerName="System.Data.SqlClient" />
    </connectionStrings>
    ```

4. Build and run the project.

## Usage

1. Open the application in your web browser.
2. Navigate to "/Import/Index".
3. Upload an Excel file with one or more worksheets.
4. Enter a table name.
5. Click the "Import" button.
6. Check the status messages for success or error information.

## Notes

- Ensure that the connection string in `Web.config` is correctly configured for your SQL Server instance.
- The application creates a table for each worksheet in the Excel file, using the provided table name as a prefix.

