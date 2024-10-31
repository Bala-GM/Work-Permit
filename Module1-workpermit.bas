Attribute VB_Name = "Module1"
Sub emailaspdf()

Dim EApp As Object
Set EApp = CreateObject("Outlook.Application")

Dim EItem As Object
Set EItem = EApp.CreateItem(0)

Dim invno As Long
Dim reqnam As String
Dim sunam As String
Dim dtsue As Date
Dim sdtsue As Date
Dim edtsue As Date
Dim path As String
Dim filnam As String
Dim nextrec As Range

invno = Range("Q4")
reqnam = Range("C11")
sunam = Range("D16")
dtsue = Range("Q2")
sdtsue = Range("D13")
edtsue = Range("D14")
path = "C:\Users\Bala Ganesh M\Desktop\VB FILES\"
filnam = invno & " - " & sunam

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, filename:=path & filnam


Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = reqnam
nextrec.Offset(0, 2) = sunam
nextrec.Offset(0, 3) = dtsue
nextrec.Offset(0, 4) = sdtsue
nextrec.Offset(0, 5) = edtsue
nextrec.Offset(0, 8) = Now

Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 6), Address:=path & filnam & ".pdf"

With EItem

    .To = Range("O15")
    .Subject = "Workorder Permit:" & invno
    .Body = " Please Find the Workorder Permit attached. "
    .Attachments.Add (path & filnam & ".pdf")
    .display

End With
    
    
End Sub





Sub Saveaspdf()

Dim invno As Long
Dim reqnam As String
Dim sunam As String
Dim dtsue As Date
Dim sdtsue As Date
Dim edtsue As Date
Dim path As String
Dim filnam As String
Dim nextrec As Range

invno = Range("Q4")
reqnam = Range("C11")
sunam = Range("D16")
dtsue = Range("Q2")
sdtsue = Range("D13")
edtsue = Range("D14")
path = "C:\Users\Bala Ganesh M\Desktop\VB FILES\"
filnam = invno & " - " & sunam

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, filename:=path & filnam

Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = reqnam
nextrec.Offset(0, 2) = sunam
nextrec.Offset(0, 3) = dtsue
nextrec.Offset(0, 4) = sdtsue
nextrec.Offset(0, 5) = edtsue

Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 6), Address:=path & filnam & ".pdf"



End Sub





Sub svasxlsx()

Dim invno As Long
Dim reqnam As String
Dim sunam As String
Dim dtsue As Date
Dim sdtsue As Date
Dim edtsue As Date
Dim path As String
Dim filnam As String
Dim nextrec As Range

invno = Range("Q4")
reqnam = Range("C11")
sunam = Range("D16")
dtsue = Range("Q2")
sdtsue = Range("D13")
edtsue = Range("D14")
path = "\\STPLPROFILE\Users\Public\share\PRINT\WorkPermit\"
filnam = invno & " - " & sunam

Sheet1.Copy

Dim shp As Shape

For Each shp In ActiveSheet.Shapes
    If shp.Type <> msoPicture Then shp.Delete
Next shp

With ActiveWorkbook
    .Sheets(1).Name = "workPermit"
    .SaveAs filename:=path & filnam, FileFormat:=51
    .Close
End With

Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = reqnam
nextrec.Offset(0, 2) = sunam
nextrec.Offset(0, 3) = dtsue
nextrec.Offset(0, 4) = sdtsue
nextrec.Offset(0, 5) = edtsue

Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 7), Address:=path & filnam & ".xlsx"

End Sub




Sub newinvoc()

Dim invno As Long

invno = Range("Q4")

MsgBox "Your next Work Permite number is " & invno + 1

Range("Q4") = invno + 1

Range("G11").Select


ThisWorkbook.Save

End Sub





Sub Record()

Dim invno As Long
Dim reqnam As String
Dim sunam As String
Dim dtsue As Date
Dim sdtsue As Date
Dim edtsue As Date
Dim nextrec As Range

invno = Range("Q4")
reqnam = Range("C11")
sunam = Range("D16")
dtsue = Range("Q2")
sdtsue = Range("D13")
edtsue = Range("D14")

Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = reqnam
nextrec.Offset(0, 2) = sunam
nextrec.Offset(0, 3) = dtsue
nextrec.Offset(0, 4) = sdtsue
nextrec.Offset(0, 5) = edtsue

End Sub
