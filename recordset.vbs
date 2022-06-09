Const adVarChar = 200  'the SQL datatype is varchar

'Create a disconnected recordset
Set rs = CreateObject("ADODB.RECORDSET")
rs.Fields.append "campo1", adVarChar, 25
rs.Fields.append "campo2", adVarChar, 25
rs.Fields.append "campo3", vbInteger, 25
rs.Fields.append "campo4", vbDate, 8

rs.CursorType = adOpenStatic
rs.Open
rs.AddNew
rs("campo1") = "Eduardo"
rs("campo2") = "Lopez"
rs("campo3") = 25
rs("campo4") = "01/01/2022"
rs.Update

rs.AddNew
rs("campo1") = "Carmelo"
rs("campo2") = "Moran"
rs("campo3") = 17
rs("campo4") = "01/10/2019"
rs.Update

rs.AddNew
rs("campo1") = "Mariana"
rs("campo2") = "Perez"
rs("campo3") = 2
rs("campo4") = "22/11/2015"
rs.Update

rs.AddNew
rs("campo1") = "Nariana"
rs("campo2") = "Pocat"
rs("campo3") = 5
rs("campo4") = "28/11/2015"
rs.Update

strList=""
rs.Sort = "campo1"
rs.MoveFirst

Do Until rs.EOF
    strList=strList & rs.Fields("campo1") & " " & rs.Fields("campo2") & " " & rs.Fields("campo3") & " " & rs.Fields("campo4") & VbCrLf
    rs.MoveNext
Loop 

'MsgBox strList
rs.MoveFirst

'Escribe en archivo (método 1)
Const ForReading = 1, ForWriting = 2,  adClipString = 2
Dim fso, f
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile("salida.txt", ForWriting, True)
 f.WriteLine "Nombre|Apellido|Num|Fecha"
 'f.Write rs.GetString(adClipString,,"|")
Do Until rs.EOF
   f.WriteLine trim(rs.Fields("campo1")) & "|" & _ 
               trim(rs.Fields("campo2")) & "|" & _ 
			   trim(rs.Fields("campo3")) & "|" & _
			   trim(rs.Fields("campo4"))
   rs.MoveNext
Loop 
 f.Close
'rs.Close
 rs.MoveFirst

'Escribe en archivo (método 2)
FileName = "salida2.txt"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile(FileName, ForWriting, True)
With rs
    While Not .EOF
        For a = 0 To .Fields.Count - 1
            oFile.Write .Fields(a).Value & "|"
        Next
		oFile.WriteLine
        .MoveNext
    Wend
End With
oFile.Close
Set oFSO = Nothing
rs.Close

'-----------------------------------
'?adBstr, vbString
' 8             8 
'?adBoolean, vbBoolean
' 11            11 
'?adInteger, vbLong
' 3             3 
'?adUnsignedTinyInt, vbByte
' 17            17 
'?adDate, vbDate
' 7             7 
'?adDouble, vbDouble
' 5             5 
'?adSingle, vbSingle
' 4             4