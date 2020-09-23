Attribute VB_Name = "Module1"
Public Sub SaveListBox(Directory As String, TheList As listbox)
    
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(savelist&)
    Next savelist&
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList As listbox)
   
    Dim MyString As String
    'On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub

Public Function FExist(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FExist = True
    Exit Function
MakeF:
        'error, file does Not exist
        FExist = False
    Exit Function
End Function
