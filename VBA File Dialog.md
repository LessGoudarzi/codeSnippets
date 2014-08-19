This routine presents a file dialog to the user to select file
---------------------------------------------------------------

``` vba

Sub getFilePointer() '
    'Declare a variable as a FileDialog object.
    Dim fd As FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd

        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the action button.
        If .Show = -1 Then

            'Step through each string in the FileDialogSelectedItems collection.
            For Each vrtSelectedItem In .SelectedItems

                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
                'MsgBox "The path is: " & vrtSelectedItem
                selectedItem = vrtSelectedItem
            Next vrtSelectedItem
        'The user pressed Cancel.
        Else
        End If
    End With

    'Set the object variable to Nothing.
    Set fd = Nothing

End Sub
```
