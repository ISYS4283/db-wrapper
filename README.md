# ISYS4283 Database Wrapper

This is a simple database abstraction layer for Microsoft SQL Server.

## Usage

By default, the connection string is set to use the ISYS4283 database.

```vb
' instantiate object
Dim db = New ISYS4283.DbWrapper.Db

' connect to database
db.ConnectionString = "Server=essql1.walton.uark.edu;Database=jeffpuckett;Trusted_Connection=yes;"

' execute DML
db.Sql = "INSERT INTO questions (question) VALUES (@question)"
db.Bind("@question", "What is a nuget package?")
db.Execute()

' fill a DataGridView
db.Sql = "SELECT * FROM questions"
db.Fill(DataGridView1)

' fill a ComboBox or ListBox (must have at least 2 columns selected)
' DisplayMember is set to first column
' ValueMember is set to second column
db.Sql = "SELECT username, id FROM users"
db.Fill(ComboBox1)
```

When using this multiple times throughout your project,
you'll want to extend the class so that your code is [DRY][dry] and
your connection string information is only in one place.

```vb
Public Class MyDb
    Inherits ISYS4283.DbWrapper.Db

    Public Sub New()
        ConnectionString = "Server=essql1.walton.uark.edu;Database=jeffpuckett;Trusted_Connection=yes;"
    End Sub
End Class
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

## Contributing

Pull requests are welcome.

### Nuget Notes

Here are some notes on building the nuget packages.

Create alias:

```
doskey pack=nuget pack ISYS4283.DbWrapper.vbproj
```

Create nuget package:

```
pack
```

Add local directory as package source in visual studio.

[nuget]:https://www.nuget.org/packages/ISYS4283.DbWrapper/
[manage-nuget]:https://i.imgur.com/20hWdUB.png
[search-isys4283]:https://i.imgur.com/2DNwZNu.png
[dry]:https://en.wikipedia.org/wiki/Don%27t_repeat_yourself
