# VBTools
A handful class modules for VB6/VBA

This stuff is already all over http://www.codereview.stackexchange.com, and as such is licensed under CC-by-SA as all Stack Exchange content is.

Enjoy!

---

###List

This class is essentially a `Collection<T>`, where all items are of type `T`... with *lots* of added functionality, largely inspired by .NET's `System.Collections.Generics.List<T>`. You'll never want to use a bare-bones `Collection` again!

###SqlCommand

Originally written for VB6 with SQL Server connections, this code works perfectly well with MySQL in VBA as well. The class can be used both as a "static class" and an object. This type/object is best used with the `UnitOfWork` class, which can encapsulate a database transaction.

###SqlResult

`SqlCommand` methods can return instead of an `ADODB.Recordset` - use it for smaller result sets, because there's a performance tradeoff here: when `SqlCommand` returns this object, the results have already been iterated once; for larger result sets it's probably better to work off the `ADODB.Recordset` directly.

###SqlResultRow

This class is essentially a *generic DTO* that `SqlResult` uses to "materialize" query results. Use the default property `Item` to refer to field names by name or by index.

###UnitOfWork

This class maintains a dictionary of `IRepository` implementations and initiates a database transaction when instantiated. The `Dispose` method rolls back any uncommitted changes, before closing the connection - this method is called automatically when the instance is terminated. Calling the `Commit` method commits the open transaction, and initiates a new one; calling the `Rollback` method rolls back the open transaction, and initiates a new one.

