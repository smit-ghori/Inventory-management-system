
# Inventory Management System

An Inventory Management System developed using ASP.NET MVC 5 and C# to manage product information with database integration.

## Overview

This project demonstrates an ASP.NET MVC web application implementing product management functionality using Entity Framework and SQL Server LocalDB.

## Features

- Product listing and management
- Create new product records
- View product details
- Edit existing product information
- Delete product records
- Server-side validation using data annotations
- Database operations using Entity Framework 6

## Technology Stack

| Technology | Details |
|---|---|
| Framework | ASP.NET MVC 5 |
| Programming Language | C# |
| ORM | Entity Framework 6 |
| Database | SQL Server LocalDB |
| Frontend | HTML, CSS, JavaScript, Razor Views |
| IDE | Visual Studio |
| Project Type | ASP.NET Web Application (.NET Framework) |

## Project Structure

```text
WebApplication1/
├── App_Start/
├── Content/
├── Controllers/
├── Models/
├── Scripts/
├── Views/
├── Web.config
├── Global.asax
└── WebApplication1.csproj
```

## Prerequisites

- Visual Studio with ASP.NET and web development support
- .NET Framework 4.7.2 targeting pack
- SQL Server LocalDB
- Git

## Installation and Setup

1. Clone the repository:

   ```bash
   git clone https://github.com/YOUR-USERNAME/YOUR-REPOSITORY.git
   ```

2. Open the `.sln` solution file in Visual Studio.

3. Restore the required NuGet packages.

4. Verify the database connection string in `Web.config`.

5. Ensure SQL Server LocalDB is installed and available.

6. Build the solution using **Build → Build Solution**.

7. Run the application using IIS Express.

## Database Configuration

The application uses Entity Framework 6 with SQL Server LocalDB.

Verify the connection string in `Web.config` before running the application. Configure the database according to the project's existing Entity Framework setup.

## Usage

1. Launch the application.
2. Navigate to the product management section.
3. Create and view product records.
4. Update product information when required.
5. Delete records when necessary.

Available operations depend on the implemented controllers and views.

## Validation and Error Handling

The application uses model validation to help ensure that submitted product information meets the defined requirements.

## Future Enhancements

- User authentication and role-based authorization
- Product search, filtering, and pagination
- Inventory quantity tracking
- Low-stock alerts
- Reporting and dashboard functionality
- Automated testing

## Limitations

This project is intended for learning and demonstration purposes. Production deployment may require additional security controls, logging, testing, and database configuration.

## Author

**Smit Ghori**

- GitHub: [smit-ghori](https://github.com/smit-ghori)
- LinkedIn: [Smit Ghori](https://www.linkedin.com/in/smit-ghori/)

## License

No license has been specified. All rights are reserved by the author unless a license is added to this repository.
