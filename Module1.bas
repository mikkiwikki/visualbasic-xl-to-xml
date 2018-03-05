Attribute VB_Name = "Module1"
Sub xl_to_xml()
' create an XML file from an Excel table
Dim MyRow As Integer, MyCol As Integer, Temp As String, YesNo As Variant, DefFolder As String
Dim XMLFileName As String, XMLRecSetName As String, MyLF As String, RTC1 As Integer
Dim RangeOne As String, RangeTwo As String, Tt As String, vchr_fld(99) As String

MyLF = Chr(10) & Chr(13)    ' a line feed command
DefFolder = "C:\"   'change this to the location of saved XML files

YesNo = MsgBox("This procedure requires the following:" & MyLF _
 & "A cellrange containing your data table" & MyLF _
 & "Are you ready to proceed?", vbQuestion + vbYesNo, "MakeXML")
 
  '& "1 A cellrange containing fieldnames (column titles)" & MyLF _
  '& "1 A filename for the XML file" & MyLF _
  '& "2 A groupname for an XML record" & MyLF _

If YesNo = vbNo Then
 Debug.Print "User aborted with 'No'"
 Exit Sub
End If

XMLFileName = FillSpaces("xl_xml_data")
'FillSpaces (InputBox("1. Enter the name of the XML file:", "MakeXML", "xl_xml_data"))
If Right(XMLFileName, 4) <> ".xml" Then
 XMLFileName = XMLFileName & ".xml"
End If

XMLRecSetName = FillSpaces("VOUCHER")
XMLRecSetName2 = FillSpaces("LINE")
XMLRecSetName3 = FillSpaces("DISTRIBUTION")

'FillSpaces (InputBox("2. Enter an identifying name of a record:", "MakeXML", "record"))

RangeOne = "A5:AF5"
'InputBox("1. Enter the range of cells containing the field names (or column titles):", "MakeXML", "A5:D5")
'If MyRng(RangeOne, 1) <> MyRng(RangeOne, 2) Then
'  MsgBox "Error: names must be on a single row" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML"
'  Exit Sub
'End If
MyRow = MyRng(RangeOne, 1)

For MyCol = MyRng(RangeOne, 3) To MyRng(RangeOne, 4)
 If Len(Cells(MyRow, MyCol).Value) = 0 Then
  MsgBox "Error: names range contains blank cell" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML"
  Exit Sub
 End If
 vchr_fld(0) = "BUSINESS_UNIT"
 vchr_fld(1) = "VOUCHER_STYLE"
 vchr_fld(2) = "INVOICE_ID"
 vchr_fld(3) = "INVOICE_DT"
 vchr_fld(4) = "VENDOR_SETID"
 vchr_fld(5) = "VENDOR_ID"
 vchr_fld(6) = "LOCATION"
 vchr_fld(7) = "ORIGIN"
 vchr_fld(8) = "GROSS_AMT"
 vchr_fld(9) = "DESCR254_MIXED"
 vchr_fld(10) = "PYMNT_HANDLING_CD"
 vchr_fld(11) = "VCHR_SRC"
 
 vchr_fld(12) = "DESCR"
 vchr_fld(13) = "QTY_VCHR"
 vchr_fld(14) = "UNIT_OF_MEASURE"
 vchr_fld(15) = "UNIT_PRICE"
 vchr_fld(16) = "MERCHANDISE_AMT"
 vchr_fld(17) = "SHIPTO_ID"
 vchr_fld(18) = "DISTRIB_MTHD_FLG"
 
 vchr_fld(19) = "DIST_AMT"
 vchr_fld(20) = "ACCOUNT"
 vchr_fld(21) = "OPERATING_UNIT"
 vchr_fld(22) = "DEPTID"
 vchr_fld(23) = "FUND"
 vchr_fld(24) = "SOURCE"
 vchr_fld(25) = "FUNCTION"
 vchr_fld(26) = "BUSINESS_UNIT_PC"
 vchr_fld(27) = "PROJECT"
 vchr_fld(28) = "ACTIVITY"
 vchr_fld(29) = "PROGRAM"
 vchr_fld(30) = "PURPOSE"
 vchr_fld(31) = "PROPERTY"
 'vchr_fld(MyCol - MyRng(RangeOne, 3)) = FillSpaces(Cells(MyRow, MyCol).Value)

Next MyCol

RangeTwo = InputBox("2. Enter the range of cells containing the data table:", "MakeXML", "A6:AF10")
If MyRng(RangeOne, 4) - MyRng(RangeOne, 3) <> MyRng(RangeTwo, 4) - MyRng(RangeTwo, 3) Then
  MsgBox "Error: number of field names <> data columns" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML"
  Exit Sub
End If

RTC1 = MyRng(RangeTwo, 3)

If InStr(1, XMLFileName, ":\") = 0 Then
 XMLFileName = DefFolder & XMLFileName
End If

Open XMLFileName For Output As #1
Print #1, "<?xml version=" & Chr(34) & "1.0" & """?>"
Print #1, "<!DOCTYPE VOUCHERS SYSTEM " & Chr(34) & "voucher.dtd" & Chr(34); "> "
Print #1, "<VOUCHERS>"

For MyRow = MyRng(RangeTwo, 1) To MyRng(RangeTwo, 2)
Print #1, "<" & XMLRecSetName & ">"
  For MyCol = RTC1 To MyRng(RangeTwo, 4)
  ' the next line uses the FormChk function to format dates and numbers
    Print #1, "<" & vchr_fld(MyCol - RTC1) & ">" & RemoveAmpersands(FormChk(MyRow, MyCol)) & "</" & vchr_fld(MyCol - RTC1) & ">"
     
    If MyCol = 12 Then
     Print #1, "<" & XMLRecSetName2 & ">"
    End If
    
    If MyCol = 19 Then
     Print #1, "<"; XMLRecSetName3; ">"
    End If
    
   Next MyCol
   Print #1, "</" & XMLRecSetName3 & ">"
  Print #1, "</" & XMLRecSetName2 & ">"
 Print #1, "</" & XMLRecSetName & ">"

Next MyRow
Print #1, "</VOUCHERS>"
Close #1
MsgBox XMLFileName & " created." & MyLF & "Process finished", vbOKOnly + vbInformation, "MakeXML"
Debug.Print XMLFileName & " saved"

End Sub
Function MyRng(MyRangeAsText As String, MyItem As Integer) As Integer
' analyse a range, where MyItem represents 1=TR, 2=BR, 3=LHC, 4=RHC

Dim UserRange As Range
Set UserRange = Range(MyRangeAsText)
Select Case MyItem
 Case 1
 MyRng = UserRange.Row
 Case 2
 MyRng = UserRange.Row + UserRange.Rows.Count - 1
 Case 3
 MyRng = UserRange.Column
 Case 4
 MyRng = UserRange.Columns(UserRange.Columns.Count).Column
End Select
Exit Function

End Function
Function FillSpaces(AnyStr As String) As String
' remove any spaces and replace with underscore character
Dim MyPos As Integer
MyPos = InStr(1, AnyStr, " ")
Do While MyPos > 0
 Mid(AnyStr, MyPos, 1) = "_"
 MyPos = InStr(1, AnyStr, " ")
Loop
FillSpaces = UCase(AnyStr)
End Function

Function FormChk(RowNum As Integer, ColNum As Integer) As String
' formats numeric and date cell values to without comma 000's and DD MMM YY
FormChk = Cells(RowNum, ColNum).Value
'If IsNumeric(Cells(RowNum, ColNum).Value) Then
 'FormChk = Format(Cells(RowNum, ColNum).Value, "#####.## ;(####.##)")
'End If
If IsDate(Cells(RowNum, ColNum).Value) Then
 FormChk = Format(Cells(RowNum, ColNum).Value, "dd mmm yy")
End If

'Data Validation -No Thrown errors, for range specified ...fix dumb user omissions.
'Spreadsheet will show blanks ...fields will filled in with these defaults.
'BUSINESS_UNIT/SETID
 If ColNum = 1 Then
  If FormChk <> "AP001" Then
   FormChk = "AP001"
  End If
 End If
 
'VOUCHER_STYLE
 If ColNum = 2 Then
  If FormChk <> "REG" Then
   FormChk = "REG"
  End If
 End If
 
'INVOICE_ID
 
'INVOICE_DT
 If ColNum = 4 Then
  FormChk = ReplaceCharacters(FormChk, "/", "")
 End If
 
'VENDOR_SETID
 If ColNum = 5 Then
  If FormChk <> "SHARE" Then
   FormChk = "SHARE"
  End If
 End If
 
'VENDOR_ID
 If ColNum = 6 Then
  If Len(FormChk) < 10 Then
   amtPad = 10 - Len(FormChk)
   FormChk = CStr(FormChk)
   For x = 1 To amtPad
      FormChk = "0" + FormChk
   Next
   Else
      'FormChk = DecToBin(txtDec(i))
   End If
 End If
 
'LOCATION
 If ColNum = 7 Then
  If FormChk <> "MAIN" Then
   FormChk = "MAIN"
  End If
 End If
 
'ORIGIN
 If ColNum = 8 Then
  If FormChk <> "ONL" Then
   FormChk = "ONL"
  End If
 End If
 
 '9 GROSS_AMT - 9 no check
 '10 descr254_mixed - no check
 
'PYMNT_HANDLING_CD
 If ColNum = 11 Then
  If FormChk <> "RE" Then
   FormChk = "RE"
  End If
 End If
 
'VCHR_SRC
 If ColNum = 12 Then
  If FormChk <> "XML" Then
   FormChk = "XML"
  End If
 End If

'DESCR 13

'QTY_VCHR
 If ColNum = 14 Then
  If FormChk = "" Then
   FormChk = "1"
  End If
 End If

'UOM
 If ColNum = 15 Then
  If FormChk = "" Then
   FormChk = "EA"
  End If
 End If
 
'UNIT_PRICE 16 - make equal to GrossAmt
 If ColNum = 16 Then
  If FormChk = "" Then
   FormChk = Format(Cells(RowNum, 9).Value)
  End If
 End If
 
'MERCHANDISE_AMT 17  - make equal to GrossAmt
 If ColNum = 17 Then
  If FormChk = "" Then
   FormChk = Format(Cells(RowNum, 9).Value)
  End If
 End If
 
'SHIP TO
 If ColNum = 18 Then
  If FormChk = "" Then
   FormChk = "0000000201"
  End If
 End If
 
'DISTRIB_MTHD_FLG
 If ColNum = 19 Then
  If FormChk = "" Then
   FormChk = "A"
  End If
 End If

'DIST_AMT
 If ColNum = 20 Then
  If FormChk = "" Then
   FormChk = Format(Cells(RowNum, 9).Value)
  End If
 End If
 
'ACCOUNT
 If ColNum = 21 Then
  If FormChk = "" Then
   FormChk = "00000"
  End If
 End If
 
 'OU
 If ColNum = 22 Then
  If FormChk = "" Then
   FormChk = "00"
  End If
 End If
 
 'Function
 If ColNum = 21 Then
  If FormChk = "" Then
   FormChk = "000"
  End If
 End If
 
'Deptid
 If ColNum = 23 Then
  If FormChk = "" Then
   FormChk = "00000"
  End If
 End If
 
'Fund code
 If ColNum = 24 Then
  If FormChk = "" Then
   FormChk = "000"
  End If
 End If
 
'Source
 If ColNum = 25 Then
  If FormChk = "" Then
   FormChk = "000000"
  End If
 End If

'function
 If ColNum = 26 Then
  If FormChk = "" Then
   FormChk = "000"
  End If
 End If

'BUPC
'PROJECT
'ACTIVITY

'program
 If ColNum = 30 Then
  If FormChk = "" Then
   FormChk = "0000"
  End If
 End If
 
'purpose
 If ColNum = 31 Then
  If FormChk = "" Then
   FormChk = "0000"
  End If
 End If
 
'property
 If ColNum = 32 Then
  If FormChk = "" Then
   FormChk = "0000"
  End If
 End If
 
End Function
Public Function ReplaceCharacters(ByRef strText As String, ByRef strUnwanted As String, ByRef strRepl As String) As String
Dim i As Integer
Dim ch As String
    For i = 1 To Len(strUnwanted)  ' Replace the unwanted character.
        strText = Replace(strText, Mid$(strUnwanted, i, 1), strRepl)
    Next
    ReplaceCharacters = strText
End Function
 
 '# CONSTANTS
'$VendorSetId = "<VENDOR_SETID>SHARE</VENDOR_SETID>\n";
'$Quantity = "<QTY_VCHR>1</QTY_VCHR>\n";
'$UnitOfMeasure = "<UNIT_OF_MEASURE>EA<UNIT_OF_MEASURE>\n";
'$Descr254Mixed = "<DESCR254_MIXED>Bailey/Howe Library Voucher</DESCR254_MIXED>\n";
'#$ShipToId = "FIXED";
'$VchrSrc = "<VCHR_SRC>XML</VCHR_SRC>\n";
'$DistMethodFlag = "<DISTRIB_MTHD_FLG>A</DISTRIB_MTHD_FLG>\n";
'$Acct = "<ACCOUNT>64501</ACCOUNT>\n";
'$OpUnit = "<OPERATING_UNIT>01</OPERATING_UNIT>\n";
'$Function = "<FUNCTION>511</FUNCTION>\n";
'$BusinessUnit = "<BUSINESS_UNIT_PC></BUSINESS_UNIT_PC>\n";
'$Project = "<PROJECT></PROJECT>\n";
'$Activity = "<ACTIVITY></ACTIVITY>\n";
'$Program = "<PROGRAM>0000</PROGRAM>\n";
'#$Purpose = "<PURPOSE>0000</PURPOSE>\n";
'$Purpose = "<PURPOSE>0495</PURPOSE>\n";
'$Property = "<PROPERTY>0000</PROPERTY>\n";
'$ShipToID = "<SHIPTO_ID>0000000005</SHIPTO_ID>\n";
'End If


Function RemoveAmpersands(AnyStr As String) As String
Dim MyPos As Integer
' replace Ampersands (&) with plus symbols (+)

MyPos = InStr(1, AnyStr, "&")
Do While MyPos > 0
 Mid(AnyStr, MyPos, 1) = "+"
 MyPos = InStr(1, AnyStr, "&")
Loop
 RemoveAmpersands = AnyStr
End Function

Function ReformatDate(AnyStr As String) As String
Dim MyPos As Integer
' Dates must be sent at MMDDYYYY, remove commas

MyPos = InStr(1, AnyStr, ",")
Do While MyPos > 0
 NewDate = Format(AnyStr, "mmddyyyy")
Loop
 ReformatDate = AnyStr
End Function

Sub Mail_Range()
'Working in 2000-2007
    Dim Source As Range
    Dim Dest As Workbook
    Dim wb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long

    Set Source = Nothing
    On Error Resume Next
    Set Source = Range("A4:AF99").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Source Is Nothing Then
        MsgBox "The source is not a range or the sheet is protected, " & _
               "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set wb = ActiveWorkbook
    Set Dest = Workbooks.Add(xlWBATWorksheet)

    Source.Copy
    With Dest.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial Paste:=xlPasteValues
        .Cells(1).PasteSpecial Paste:=xlPasteFormats
        .Cells(1).Select
        Application.CutCopyMode = False
    End With

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Selection of " & wb.Name & " " _
                 & Format(Now, "dd-mmm-yy h-mm-ss")

    If Val(Application.Version) < 12 Then
        'You use Excel 2000-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        'You use Excel 2007
        FileExtStr = ".xlsx": FileFormatNum = 51
    End If

    With Dest
        .SaveAs TempFilePath & TempFileName & FileExtStr, _
                FileFormat:=FileFormatNum
        On Error Resume Next
        .SendMail "Monica.Devino@uvm.edu", _
                  "Voucher File for Procurement Review"
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    Kill TempFilePath & TempFileName & FileExtStr

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

