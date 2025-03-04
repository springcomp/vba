# Writing Better Code

## Option Explicit

Your first sample code is not using good practices as it uses implicit variables and does not specify the type of variables. This prevents the
compiler from checking the code is valid.

This will also help you write code as the compiler will provide _Intellisense_ when you use object methods and properties.

Add the following line to the beginning of the module.

```basic
Option Explicit

```

This instructs the compiler that all variables must be declared before use. In the menu, select the **Debug|Compile VBA Project**. Notice that
the code is now considered not valid as you must declare the variables and
their types.

Change the module with the following more correct code below:

```basic
Public Sub Sample()

    Dim nCount As Long
    Dim nIndex As Integer
    
    Let nCount = ThisWorkbook.Worksheets.Count
    Let nIndex = 1
    For nIndex = 1 To nCount
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Item(nIndex)
        MsgBox ws.Name
    Next
    
    Set ws = Nothing

End Sub
```

Alternatively, using the **Import File…** option to import the code from the [03-modSample.bas](modules/03-modSample.bas) file. You may need to rename the module and/or remove the previous copy of that module.

## Assignment

In the previous section, you learned about _Scalars_ vs _Objects_.

When working with _Functions_ you need to assign their return values to your variables. The assignment instruction depends on the value type. A scalar value can be assigned using the `Let` keyword.

```basic
Dim nIndex As Integer ' a scalar
Let nIndex = 1 ' use the Let keyword (good practice)
nIndex = 1 ' the Let keyword is optional
```

An object value must be assigned using the `Set` keyword.

```basic
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets.Item(nIndex) ' assign object
```

Likewise, if you create a _Function_ that returns a scalar,
you may use the `Let` keyword to set the return value.

```basic
Public Function MyScalarFunction() As String ' a scalar type
  Let MyScalarFunction = "forty-two"
End Function

However, if the _Function_ returns an object,
you must use the `Set` keyword to set the return value.

```basic
Public Function FirstWorksheet() As Worksheet ' an object type
  Set FirstWorksheet = ThisWorkbook.Worksheets.Index(1)
End Function
```

In the next section, you will learn to create your own objects.