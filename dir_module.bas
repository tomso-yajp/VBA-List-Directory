Attribute VB_Name = "dir_module"
Const com As String = ","
Dim fol As Variant, fil As Variant

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  debug_mail : call debug                                                    XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub debug_main()
Dim p As String
'make folders and files
p = new_dir

'display folder names as a list in the debug window
fol = "": Call dir_fol(p)
Call output_debug(Mid(fol, 1, Len(fol) - 1), "dir_fol")

'display file names as a list in the debug window
fol = "": fil = "": Call dir_fil(p)
Call output_debug(Mid(fil, 1, Len(fil) - 1), "dir_fil")

'display folder names and file names as a list in the debug window
Call dir_main
Call output_debug(fol, "dir_sub:folder")
Call output_debug(fil, "dir_sub:file")

'display file names as a list in the sheet
Call output_sheet(fil, "file_list")

'display folder names as a list in the sheet
Call output_sheet(fol, "folder_list")

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  dir_fol : get the folder pash                                              XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function dir_fol(ByVal p As String)
Dim fo As Variant
With CreateObject("scripting.filesystemobject")
  For Each fo In .getfolder(p).subfolders
    fol = fol & fo & com
    Call dir_fol(CStr(fo))
  Next
End With
dir_fol = Mid(fol, 1, Len(fol) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  dir_fil : get the file pash                                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function dir_fil(ByVal p As String)
Dim fi As Variant, f As Variant
Dim i As Integer:
Call dir_fol(p): f = Split(Mid(fol, 1, Len(fol) - 1), com)
With CreateObject("scripting.filesystemobject")
  For i = UBound(f) To 0 Step -1
    For Each fi In .getfolder(f(i)).Files
      If InStr(fil & com, fi & com) = 0 Then _
        fil = fi & com & fil
    Next
  Next
End With
f = Split(fil, com)
dir_fil = Mid(fil, 1, Len(fil) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  dir_main : call dir_sub                                                    XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub dir_main()
Dim p As String
  fol = "": fil = ""
  p = new_dir
  Call dir_sub(p)
  fol = Mid(fol, 1, Len(fol) - 1)
  fil = Mid(fil, 1, Len(fil) - 1)
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  dir_sub : get the pash. path is folders and files                          XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub dir_sub(Optional p As String)
Dim fo As Variant, fi As Variant
With CreateObject("scripting.filesystemobject")
  For Each fo In .getfolder(p).subfolders
    fol = fol & fo & com
    Call dir_sub(CStr(fo))
  Next
  For Each fi In .getfolder(p).Files
    fil = fil & fi & com
  Next
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  new_dir : make folders and files                                           XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function new_dir()
Dim p As String: p = ThisWorkbook.Path & "\test"
Dim i As Integer, bl As Integer: bl = 0
Dim fname As String
With CreateObject("scripting.filesystemobject")
  If dir(p, vbDirectory) <> "" Then Call del_dir(p)
  MkDir p
  For i = 1 To 9
    fname = p & "\" & i: MkDir fname
    .CreateTextFile fname & "\" & i & ".txt", True
    If i Mod 2 = 0 Then
      fname = fname & "\" & i & "_1": MkDir fname
      .CreateTextFile fname & "\" & i & "_1.txt", True
    End If
  Next
  new_dir = p
End With
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  del_dir : delete folders and files                                         XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub del_dir(ByVal p As String)
Dim fo As Variant, fi As Variant
With CreateObject("scripting.filesystemobject")
  For Each fo In .getfolder(p).subfolders
    Call del_dir(CStr(fo))
  Next
  For Each fi In .getfolder(p).Files
    Kill fi
  Next
  RmDir p
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  output_debug : output sheet name into debug window                         XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub output_debug(Optional v As Variant = "", Optional mname As String = "")
Dim d As Variant: d = Split(v, com)
Dim i As Integer
Debug.Print "========================================================"
Debug.Print "  call is " & mname & ". output list of dirs"
Debug.Print "--------------------------------------------------------"
For i = 0 To UBound(d)
  Debug.Print d(i)
Next
Debug.Print "========================================================"

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  output_sheet : output sheet name in sheet                                  XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub output_sheet(Optional v As Variant = "", Optional sname As String = "make")
Dim slist As String: slist = sheet_list
Dim d As Variant: d = com & v: d = Split(d, com)
Dim rw As Integer, col As Integer
Dim i As Integer, n As Integer
rw = 2: col = 1:
With ThisWorkbook
  If InStr(com & slist & com, com & sname & com) = 0 Then _
    .Worksheets.Add.Name = sname
  With .Worksheets(sname)
    .Cells.Clear
    .Cells(rw, col) = "no": .Cells(rw, col + 1) = "list"
    For i = 1 To UBound(d)
      .Cells(rw + i, col) = "=row() -" & rw: .Cells(rw + i, col + 1) = d(i)
    Next
    With .Sort
      .SortFields.Clear
      .SortFields.Add Key:=Range(Cells(rw, col + 1), Cells(UBound(d), col + 1)), _
        Order:=xlAscending
      .SetRange Range(Cells(rw, col + 1), Cells(UBound(d), col + 1))
      .Header = xlYes
      .Apply
    End With
  End With
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  sheet_list : get sheet name                                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function sheet_list()
Dim s As Variant, sname As String
With ThisWorkbook
  For Each s In .Worksheets
    sname = sname & s.Name & com
  Next
End With
sheet_list = Mid(sname, 1, Len(sname) - 1)
End Function
