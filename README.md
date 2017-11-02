# ISYS4283 Database Wrapper

This is a simple database abstraction layer for Microsoft SQL Server.

## Usage

Import the namespace at the top of your class file.

```vb
Imports ISYS4283.DbWrapper
```

By default, the connection string is set to use the ISYS4283 database.

```vb
' connect to database
db.ConnectionString = "Server=essql1.walton.uark.edu;Database=jeffpuckett;Trusted_Connection=yes;"

' execute DML
db.Sql = "INSERT INTO questions (question) VALUES (@question)"
db.Bind("@question", "What is a nuget package?")
db.Execute()

' fill a DataGridView
db.Sql = "SELECT * FROM questions"
db.Fill(DataGridView1)
```

### Exception Logging

By default, a `MsgBox` will simply show the error message.
To customize this, extend the `Db` class and override the `Log` method.

```vb
Protected Overrides Sub Log(ByRef exception As Exception)
    ' your custom logger implementation
End Sub
```

## Installation

Grab the [package from nuget][nuget].

In the solution explorer, right click your project references
and select `Manage NuGet Packages`

![manage nuget packages][manage-nuget]

On the browse tab, search for `ISYS4283.DbWrapper`
and click the install icon on the far right for the latest version.

![search and install][search-isys4283]

[nuget]:https://www.nuget.org/packages/ISYS4283.DbWrapper/
[manage-nuget]:https://i.imgur.com/20hWdUB.png
[search-isys4283]:https://i.imgur.com/2DNwZNu.png
