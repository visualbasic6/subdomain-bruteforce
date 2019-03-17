Attribute VB_Name = "sdbfMod"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Sub FlatBorder(ByVal hwnd As Long)

    Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub

Public Function LoadListFromFile(ByRef SourceFile As String, _
     ByRef ToFormList As ListBox)

    On Error GoTo ErrEvt
    Dim TextLine As String, FN As Integer

    On Error Resume Next

    FN = FreeFile
     Open SourceFile For Input As #FN
       Do While Not EOF(FN)
       Line Input #FN, TextLine
       If TextLine <> LineToRem Then
        ToFormList.AddItem (TextLine)
       End If
    Loop
    Close #FN
    
    Exit Function
ErrEvt:
    Select Case Err.Number
       Case 51
          Err.Clear
       Case Else
    End Select
    Resume Next

End Function

Public Function SaveListToFile(ByVal strPrintToFile As String, _
   ByRef lstFormList As ListBox, Optional ByVal blnClearList As Boolean = False)

    On Error Resume Next

    Dim I As Long
    Dim FN As Integer

    FN = FreeFile

  Open strPrintToFile For Output As #FN

   For I = 0 To lstFormList.ListCount - 1
      Print #FN, lstFormList.List(I)
   Next I

  Close #FN

  If blnClearList = True Then lstFormList.Clear

End Function

Public Sub SaveText(sString As String, sFile As String)
    On Error Resume Next
    Open sFile$ For Output As #1
        Print #1, sString$
    Close #1
End Sub
