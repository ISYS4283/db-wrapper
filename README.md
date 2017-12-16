# ISYS4283 Database Wrapper

This is a simple database abstraction layer for Microsoft SQL Server.

## Usage

Extend the class and set your connection string in the constructor.

### VB

```vb
Public Class Db
    Inherits ISYS4283.DbWrapper.Db

    Public Sub New()
        ConnectionString = "Server=essql1.walton.uark.edu;Database=insert_database_name_here;Trusted_Connection=yes;"
    End Sub
End Class
```

### C#

```cs
class Db : ISYS4283.DbWrapper.Db
{
    public Db()
    {
        ConnectionString = "Server=essql1.walton.uark.edu;Database=insert_database_name_here;Trusted_Connection=yes;";
    }
}
```

Here's some examples filling form controls from a database in VB.

```vb
' instantiate object
Dim db = New Db

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
