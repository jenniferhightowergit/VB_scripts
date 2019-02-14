Sub audit_IFX_to_CBASE()

'use flight specific diagram from IFX
'transfers each stowage position to output in flight analysis

Dim AC_FROM_IFX As Worksheet
Set AC_FROM_IFX = Worksheets("AC_FROM_IFX")

Dim AC_FROM_CBASE As Worksheet
Set AC_FROM_CBASE = Worksheets("AC_FROM_CBASE")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim STOWAGE_MATCH As Worksheet
Set STOWAGE_MATCH = Worksheets("STOWAGE_MATCH")



Dim i As Long
Dim k As Long
Dim j As Long
Dim w As Long
Dim y As Long
Dim TempCell As String
Dim TempCellUp As String
Dim TempCellDown As String
Dim COMMA_POS As Integer
Dim var1 As String
Dim var2 As String



w = 1

AC_FROM_IFX.Select
AC_FROM_IFX.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

'fill in stowage positions in IFX diagram

For i = toprow1 To endrow1

    TempCell = AC_FROM_IFX.Cells(i, 37).Value
    TempCellUp = AC_FROM_IFX.Cells(i - 1, 37).Value
    TempCellDown = AC_FROM_IFX.Cells(i + 1, 37).Value
    
    If TempCell <> "" And TempCellDown = "" Then
    
        AC_FROM_IFX.Cells(i + 1, 37).Value = TempCell
    
    End If
    
    If TempCell = "" And TempCellUp <> "" Then
    
        AC_FROM_IFX.Cells(i, 37).Value = TempCellUp
    
    End If
    
    If TempCellDown <> "" And TempCellDown <> TempCell Then
    
        i = i + 1
    
    End If

Next i

Dim COUNTER As Integer
Dim STOWAGE_LEN As Integer


AC_FROM_IFX.Select
AC_FROM_IFX.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

w = 1
OUTPUT.Cells(w, 2).Value = "AIR CRAFT"
OUTPUT.Cells(w, 1).Value = "POSITION"
OUTPUT.Cells(w, 3).Value = "PART NUMBER"
OUTPUT.Cells(w, 4).Value = "PART NUMBER DESCRIPTION"
w = w + 1

'transfer position, position part number, partnumber name, if drawer give customer number, drawer description


For i = toprow1 To endrow1

    OUTPUT.Cells(w, 1).Value = AC_FROM_IFX.Cells(i, 37).Value
        
    If AC_FROM_IFX.Cells(i, 22).Value = 1 Then
        'drawer part number and drawer name
        OUTPUT.Cells(w, 2).Value = AC_FROM_IFX.Cells(i, 2).Value
        OUTPUT.Cells(w, 3).Value = AC_FROM_IFX.Cells(i, 24).Value
        OUTPUT.Cells(w, 4).Value = AC_FROM_IFX.Cells(i, 18).Value
        OUTPUT.Cells(w, 7).Value = AC_FROM_IFX.Cells(i, 14).Value
        w = w + 1
        
    Else
        'part number and part name
        OUTPUT.Cells(w, 2).Value = AC_FROM_IFX.Cells(i, 2).Value
        OUTPUT.Cells(w, 3).Value = AC_FROM_IFX.Cells(i, 24).Value
        OUTPUT.Cells(w, 4).Value = AC_FROM_IFX.Cells(i, 25).Value
        OUTPUT.Cells(w, 7).Value = AC_FROM_IFX.Cells(i, 14).Value
        w = w + 1
    
    End If



Next i


OUTPUT.Select
OUTPUT.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

'go through position and if there are multiple seperate them out by a comma


For i = toprow1 To 3000

    If InStr(1, OUTPUT.Cells(i, 1).Value, ",") <> 0 Then
    
    'find a comma at all
    

    For j = 1 To Len(OUTPUT.Cells(i, 1).Value)
    
    
    'MsgBox (j)
    'MsgBox (OUTPUT.Cells(i, 1).Value)
    
    'find the first comma attach first value to a variable and the rest to a seperate variable,
    'create a duplicate line, use first variable for first line then second variable for second line
        
        
        'stops loop once the line no longer has a comma
        If InStr(1, OUTPUT.Cells(i, 1).Value, ",") <> 0 Then
        
        COMMA_POS = InStr(1, OUTPUT.Cells(i, 1).Value, ",")
        'MsgBox (COMMA_POS)
        
        'create first and second variable
        var1 = Mid(OUTPUT.Cells(i, 1).Value, 1, COMMA_POS - 1)
        var2 = Mid(OUTPUT.Cells(i, 1).Value, COMMA_POS + 1, Len(OUTPUT.Cells(i, 1).Value))
        'MsgBox (var1)
        'MsgBox (var2)
            
        'insert line
        Range(OUTPUT.Cells(i, 1), OUTPUT.Cells(i, 4)).Select
        Selection.Copy
        Range(OUTPUT.Cells(i + 1, 1), OUTPUT.Cells(i + 1, 10)).Select
        Selection.Insert Shift:=xlDown
        
        endrow1 = endrow1 + 1
        
        'change values of first and second line
        
        OUTPUT.Cells(i, 1).Value = var1
        OUTPUT.Cells(i + 1, 1).Value = var2
        i = i + 1
        
        Else
        Exit For
        
        End If
        
        
        'MsgBox (NEXT_START_POS)
        'MsgBox (AC_FROM_IFX.Cells(i, 37).Value)
        'MsgBox (j)
        
    Next j

    'if no comma just lists the stowage position assuming it is a singular position
   
    End If
    'MsgBox (AC_FROM_IFX.Cells(i, 37).Value)


Next i


'check IFX stowage against CBASE stowage and pull the cbase name and code

OUTPUT.Select
OUTPUT.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

STOWAGE_MATCH.Select
STOWAGE_MATCH.Cells(1, 1).Select
toprow2 = 2
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row
Selection.End(xlUp).Select

AC_FROM_CBASE.Select
AC_FROM_CBASE.Cells(1, 1).Select
toprow3 = 2
Selection.End(xlDown).Select
endrow3 = ActiveCell.Row
Selection.End(xlUp).Select


For i = toprow1 To endrow1

    For j = toprow2 To endrow2
    
        If OUTPUT.Cells(i, 2).Value = STOWAGE_MATCH.Cells(j, 1).Value And OUTPUT.Cells(i, 1).Value = STOWAGE_MATCH.Cells(j, 2).Value Then
        
            For k = toprow3 To endrow3
            
                If STOWAGE_MATCH.Cells(j, 3).Value = AC_FROM_CBASE.Cells(k, 1).Value Then
                
                    OUTPUT.Cells(i, 5).Value = AC_FROM_CBASE.Cells(k, 5).Value
                    OUTPUT.Cells(i, 6).Value = AC_FROM_CBASE.Cells(k, 4).Value
                
                End If
            
            Next k
    
        End If
    
    Next j




Next i




End Sub


Sub doubles()

'show doulbes second leg flight with mileage
'with CFS - make sur eto remove dl first!!!!!!

Dim citypairs_mileage As Worksheet
Set citypairs_mileage = Worksheets("citypairs_mileage")

Dim CFS As Worksheet
Set CFS = Worksheets("CFS")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")


Dim i As Long
Dim k As Long
Dim j As Long

CFS.Select
CFS.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select


CFS.Cells(1, 11).Value = "CONCAT"

For i = 2 To endrow1

    CFS.Cells(i, 11).Value = CStr(CFS.Cells(i, 1).Value) + CStr(CFS.Cells(i, 2).Value) + CStr(CFS.Cells(i, 5).Value)

Next i


CFS.Select
CFS.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

OUTPUT.Select
OUTPUT.Cells(1, 1).Select
toprow2 = 2
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row
Selection.End(xlUp).Select

citypairs_mileage.Select
citypairs_mileage.Cells(1, 1).Select
toprow3 = 2
Selection.End(xlDown).Select
endrow3 = ActiveCell.Row
Selection.End(xlUp).Select


For i = 2 To endrow1

    For j = 2 To endrow3
    
        If OUTPUT.Cells(i, 8).Value = citypairs_mileage.Cells(j, 1).Value And OUTPUT.Cells(i, 9).Value = citypairs_mileage.Cells(j, 2).Value Then
        
            OUTPUT.Cells(i, 11).Value = citypairs_mileage.Cells(j, 4).Value
        
        End If
    
    
    Next j


Next i


CFS.Select
CFS.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

OUTPUT.Select
OUTPUT.Cells(1, 1).Select
toprow2 = 2
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row
Selection.End(xlUp).Select


citypairs_mileage.Select
citypairs_mileage.Cells(1, 1).Select
toprow3 = 2
Selection.End(xlDown).Select
endrow3 = ActiveCell.Row
Selection.End(xlUp).Select

'add second leg first

OUTPUT.Cells(1, 10).Value = "SECOND LEG CITY"
OUTPUT.Cells(1, 11).Value = "MILEAGE LEG 1"
OUTPUT.Cells(1, 12).Value = "MILEAGE LEG 2"

k = 2

For i = 2 To endrow1

For j = 2 To endrow2

    If CFS.Cells(i, 11).Value = OUTPUT.Cells(j, 1).Value And CFS.Cells(i + 1, 7).Value = 2 Then
    
        OUTPUT.Cells(j, 10).Value = CFS.Cells(i + 1, 6).Value
           
    End If
    
Next j
    
Next i

'add mileage

OUTPUT.Select
OUTPUT.Cells(1, 1).Select
toprow2 = 2
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row
Selection.End(xlUp).Select


citypairs_mileage.Select
citypairs_mileage.Cells(1, 1).Select
toprow3 = 2
Selection.End(xlDown).Select
endrow3 = ActiveCell.Row
Selection.End(xlUp).Select

For i = 2 To endrow2

    For j = 2 To endrow3
        
            If OUTPUT.Cells(i, 9).Value = citypairs_mileage.Cells(j, 1).Value And OUTPUT.Cells(i, 10).Value = citypairs_mileage.Cells(j, 2).Value Then
        
                OUTPUT.Cells(i, 12).Value = citypairs_mileage.Cells(j, 4).Value
        
            End If
         
    Next j

Next i

End Sub



Sub MergeLoadingLists()

Dim Flight_Overview As Worksheet
Set Flight_Overview = Worksheets("Flight_Overview")

Dim i As Long
Dim CONCAT As String
Dim endrow1 As Long
Dim toprow1 As Long



Flight_Overview.Select
Flight_Overview.Cells(1, 1).Select
toprow1 = 2
endrow1 = 30000


For i = toprow1 To endrow1

    CONCAT = Flight_Overview.Cells(i, 15).Value
    
    If Flight_Overview.Cells(i, 16).Value <> "" Then
    
        CONCAT = CONCAT + "/" + Flight_Overview.Cells(i, 16).Value
    
    End If
    
    If Flight_Overview.Cells(i, 17).Value <> "" Then
    
        CONCAT = CONCAT + "/" + Flight_Overview.Cells(i, 17).Value
    
    End If
    
    If Flight_Overview.Cells(i, 18).Value <> "" Then
    
        CONCAT = CONCAT + "/" + Flight_Overview.Cells(i, 18).Value
    
    End If
    
    If Flight_Overview.Cells(i, 19).Value <> "" Then
    
        CONCAT = CONCAT + "/" + Flight_Overview.Cells(i, 19).Value
    
    End If
    
    
    
    
    Flight_Overview.Cells(i, 14).Value = CONCAT

Next i




End Sub





Sub ScheduledCodes_with_LoadingLists()

Dim ScheduledCodes As Worksheet
Set ScheduledCodes = Worksheets("ScheduledCodes")

Dim Flight_Overview As Worksheet
Set Flight_Overview = Worksheets("Flight_Overview")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim i As Long
Dim j As Long
Dim k As Long

k = 1

'compare IFX scheduled codes against Cbase first leg flight so you can catch potentially missing flights

Dim CONCAT As String


Flight_Overview.Select
Flight_Overview.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

Flight_Overview.Cells(1, 9).Value = "CONCAT"

For i = 2 To endrow1

    Flight_Overview.Cells(i, 9).Value = CStr(Flight_Overview.Cells(i, 2).Value) + CStr(Flight_Overview.Cells(i, 1).Value) + CStr(Flight_Overview.Cells(i, 7).Value)


Next i



ScheduledCodes.Cells(1, 1).Value = "CONCAT"

For i = 2 To 3000

    ScheduledCodes.Cells(i, 1).Value = CStr(ScheduledCodes.Cells(i, 2).Value) + CStr(ScheduledCodes.Cells(i, 10).Value) + CStr(ScheduledCodes.Cells(i, 3).Value)


Next i


Flight_Overview.Select
Flight_Overview.Cells(1, 1).Select
toprow1 = 2
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select

ScheduledCodes.Select
ScheduledCodes.Cells(1, 1).Select
toprow2 = 2

endrow2 = 3000

        OUTPUT.Cells(k, 1).Value = "CONCAT"
        OUTPUT.Cells(k, 2).Value = "Flight Number"
        OUTPUT.Cells(k, 3).Value = "Scheduled Codes"
        OUTPUT.Cells(k, 4).Value = "Loading"
        OUTPUT.Cells(k, 5).Value = "Departure Time"
        OUTPUT.Cells(k, 6).Value = "Departure Date"
        OUTPUT.Cells(k, 7).Value = "Aircraft"
        OUTPUT.Cells(k, 8).Value = "Departure Station"
        OUTPUT.Cells(k, 9).Value = "Arrival Station"

For i = toprow1 To endrow1

   For j = toprow2 To endrow2
   
    If ScheduledCodes.Cells(j, 1).Value = Flight_Overview.Cells(i, 9).Value Then
        k = k + 1
        OUTPUT.Cells(k, 1).Value = ScheduledCodes.Cells(j, 1).Value
        OUTPUT.Cells(k, 2).Value = ScheduledCodes.Cells(j, 2).Value
        OUTPUT.Cells(k, 8).Value = ScheduledCodes.Cells(j, 3).Value
        OUTPUT.Cells(k, 9).Value = ScheduledCodes.Cells(j, 6).Value
        OUTPUT.Cells(k, 5).Value = ScheduledCodes.Cells(j, 4).Value
        OUTPUT.Cells(k, 6).Value = Flight_Overview.Cells(i, 1).Value
        OUTPUT.Cells(k, 7).Value = Flight_Overview.Cells(i, 6).Value
        OUTPUT.Cells(k, 3).Value = ScheduledCodes.Cells(j, 21).Value
        OUTPUT.Cells(k, 4).Value = Flight_Overview.Cells(i, 15).Value + "/" + Flight_Overview.Cells(i, 16).Value + "/" + Flight_Overview.Cells(i, 17).Value + "/" + Flight_Overview.Cells(i, 18).Value + "/" + Flight_Overview.Cells(i, 19).Value 'loading
        
   
    End If
   
   Next j

Next i




End Sub


Sub allOneLine()

'use for IFX SCHEDULED codes

Dim ScheduledCodes As Worksheet
Set ScheduledCodes = Worksheets("ScheduledCodes")

ScheduledCodes.Cells(1, 21).Value = "CONCATENATED SCHEDULED CODES"

Dim CONCAT As Variant
Dim i As Long
Dim j As Long
Dim MarkedRow As Integer
Dim Newcolor As String




 

ScheduledCodes.Select
ScheduledCodes.Cells(1, 1).Select


toprow1 = ScheduledCodes.Cells(1, 1).Row
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row
Selection.End(xlUp).Select


'put junk on first llrelevant line

For i = 2 To endrow1
  
    
    
    
    
    'MsgBox (ScheduledCodes.Cells(i, 15).Row)
    
      If Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" And Left(ScheduledCodes.Cells(i + 1, 15).Value, 11) = "First Class" Then
      
      i = i + 1
      
      
      End If

      
      
      If Left(ScheduledCodes.Cells(i, 15).Value, 9) = "Delta One" And Left(ScheduledCodes.Cells(i + 1, 15).Value, 11) = "First Class" And ScheduledCodes.Cells(i, 2).Value = ScheduledCodes.Cells(i + 1, 2).Value Then
      
         MarkedRow = ScheduledCodes.Cells(i, 15).Row
        
        i = i + 1
        
            If Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" Then
            
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value) + ";" + "*" + CStr(ScheduledCodes.Cells(i, 18).Value) + "*" + ";"
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                    i = i + 1
                    
                    Else
                    
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value)
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                    i = i + 1
            
            End If
        
        TEMP = "FALSE"
        
        Do Until TEMP = "TRUE"
        
        
            If Left(ScheduledCodes.Cells(i, 15).Value, 7) = "Comfort" Or Left(ScheduledCodes.Cells(i, 15).Value, 10) = "Main Cabin" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "FltAtt" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "Common" Or Left(ScheduledCodes.Cells(i, 15).Value, 5) = "Pilot" Then
          
                If ScheduledCodes.Cells(i, 18).Value <> "" Then
                
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value) + ";" + "*" + CStr(ScheduledCodes.Cells(i, 18).Value) + "*" + ";"
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                    
                    Else
                    
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value)
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                
                End If
                
                
                Else
                
                If Left(ScheduledCodes.Cells(i - 1, 15).Value, 7) = "Comfort" Or Left(ScheduledCodes.Cells(i - 1, 15).Value, 10) = "Main Cabin" Or Left(ScheduledCodes.Cells(i - 1, 15).Value, 6) = "FltAtt" Or Left(ScheduledCodes.Cells(i - 1, 15).Value, 6) = "Common" Or Left(ScheduledCodes.Cells(i - 1, 15).Value, 5) = "Pilot" Then
          
                If ScheduledCodes.Cells(i - 1, 18).Value <> "" Then
                
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i - 1, 17).Value) + ";" + "*" + CStr(ScheduledCodes.Cells(i - 1, 18).Value) + "*" + ";"
                    ScheduledCodes.Cells(i - 1, 1).Select
                    Range(ScheduledCodes.Cells(i - 1, 1), ScheduledCodes.Cells(i - 1, 21)).Select
                    Selection.Delete Shift:=xlUp
                    
                    Else
                    
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i - 1, 17).Value)
                    ScheduledCodes.Cells(i - 2, 1).Select
                    Range(ScheduledCodes.Cells(i - 1, 1), ScheduledCodes.Cells(i - 1, 21)).Select
                    Selection.Delete Shift:=xlUp
                
                End If
                
          
            End If
                
          
            End If
            
            
            
            
            If Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" Or Left(ScheduledCodes.Cells(i, 15).Value, 9) = "Delta One" Or Left(ScheduledCodes.Cells(i, 15).Value, 9) = " " Then
            
                Exit Do
            
            End If
            
            If Left(ScheduledCodes.Cells(i, 15).Value, 6) = "Common" Then
            
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value)
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
            
            End If
            
        Loop
        
        ScheduledCodes.Cells(MarkedRow, 21).Value = CONCAT
        i = i - 1
          
      End If
      
      
      
      
      If Left(ScheduledCodes.Cells(i, 15).Value, 9) = "Delta One" And Left(ScheduledCodes.Cells(i + 1, 15).Value, 11) <> "First Class" Then
      
        MarkedRow = ScheduledCodes.Cells(i, 15).Row
        
        i = i + 1
        
        TEMP = "FALSE"
        
        Do Until TEMP = "TRUE"
        
        
            If Left(ScheduledCodes.Cells(i, 15).Value, 7) = "Comfort" Or ScheduledCodes.Cells(i, 15).Value = "" Or Left(ScheduledCodes.Cells(i, 15).Value, 10) = "Main Cabin" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "FltAtt" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "Common" Or Left(ScheduledCodes.Cells(i, 15).Value, 5) = "Pilot" Then
          
                If ScheduledCodes.Cells(i, 18).Value <> "" Then
                
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value) + ";" + "*" + CStr(ScheduledCodes.Cells(i, 18).Value) + "*" + ";"
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                    
                    Else
                    
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value)
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                
                End If
                
          
            End If
            
            'MsgBox (i)
            
            If Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" Or Left(ScheduledCodes.Cells(i, 15).Value, 9) = "Delta One" Or ScheduledCodes.Cells(i, 1).Value = "" Or ScheduledCodes.Cells(i, 2).Value <> ScheduledCodes.Cells(i - 1, 2).Value Then
            
                Exit Do
            
            End If
            
        Loop
        
        ScheduledCodes.Cells(MarkedRow, 21).Value = CONCAT
        i = i - 1
            
           
      End If
      
      
      
      
      
      If (Left(ScheduledCodes.Cells(i - 1, 15).Value, 9) <> "Delta One" And Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" And Left(ScheduledCodes.Cells(i + 1, 15).Value, 11) <> "First Class") Or (Left(ScheduledCodes.Cells(i - 1, 15).Value, 9) = "Delta One" And Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" And ScheduledCodes.Cells(i, 15).Value <> ScheduledCodes.Cells(i - 1, 15).Value) Then
      
         MarkedRow = ScheduledCodes.Cells(i, 15).Row
        
        i = i + 1
        
        TEMP = "FALSE"
        
        Do Until TEMP = "TRUE"
        
        
            If Left(ScheduledCodes.Cells(i, 15).Value, 7) = "Comfort" Or Left(ScheduledCodes.Cells(i, 15).Value, 10) = "Main Cabin" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "FltAtt" Or Left(ScheduledCodes.Cells(i, 15).Value, 6) = "Common" Or Left(ScheduledCodes.Cells(i, 15).Value, 5) = "Pilot" Then
          
                If ScheduledCodes.Cells(i, 18).Value <> "" Then
                
                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value) + ";" + "*" + CStr(ScheduledCodes.Cells(i, 18).Value) + "*" + ";"
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                    
                    Else

                    CONCAT = CONCAT + CStr(ScheduledCodes.Cells(i, 17).Value)
                    ScheduledCodes.Cells(i, 1).Select
                    Range(ScheduledCodes.Cells(i, 1), ScheduledCodes.Cells(i, 21)).Select
                    Selection.Delete Shift:=xlUp
                
                End If
                
          
            End If
            
            If Left(ScheduledCodes.Cells(i, 15).Value, 11) = "First Class" Or Left(ScheduledCodes.Cells(i, 15).Value, 9) = "Delta One" Then
            
                Exit Do
            
            End If
            
    
            
            If ScheduledCodes.Cells(i, 15).Value = "" Or ScheduledCodes.Cells(i, 15).Value = " " Then
            
            Exit Do
            
            End If
            
        Loop
        
        ScheduledCodes.Cells(MarkedRow, 21).Value = CONCAT
        i = i - 1
        CONCAT = ""
           
      End If
      
      


Next i




End Sub





Sub CheckifContentChanged()

'this script is to check the current diagram to the future diagram in GP4 and see if any changes have been made
' Use ife CheckMonthlyCHanges to run script

'it would be good to accend the ifx packing list list in the same order to make the comparisons faster

Dim CURRENT As Worksheet
Set CURRENT = Worksheets("CURRENT")
Dim FUTURE As Worksheet
Set FUTURE = Worksheets("FUTURE")
Dim OUTPUT_CURRENT As Worksheet
Set OUTPUT_CURRENT = Worksheets("OUTPUT_CURRENT")
Dim OUTPUT_FUTURE As Worksheet
Set OUTPUT_FUTURE = Worksheets("OUTPUT_FUTURE")
Dim DISCREPANCIES As Worksheet
Set DISCREPANCIES = Worksheets("DISCREPANCIES")


Dim i As Long
Dim j As Long
Dim x As Long
Dim y As Long

Dim CONCAT As Variant


'stowage list, F,K,N,U,V,Z,AA,AB from current


CURRENT.Select
CURRENT.Cells(1, 1).Select
toprow1 = CURRENT.Cells(1, 1).Row
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row

Dim p As Long
Dim rows As String
Dim rang As Variant
Dim AC As String
Dim F As String
Dim K2 As String
Dim n As String
Dim U As String
Dim V As String
Dim Z As String
Dim AA As String
Dim AB As String

Dim TempArray As String




'current data
OUTPUT_CURRENT.Select

k = 1
p = 1

For i = 2 To endrow1

    
    If CURRENT.Cells(i, 34).Value <> "" Then
    
        AC = CStr(CURRENT.Cells(i, 1).Value)
        F = CURRENT.Cells(i, 6).Value
        K2 = CURRENT.Cells(i, 11).Value
        n = CURRENT.Cells(i, 14).Value
        U = CURRENT.Cells(i, 21).Value
        V = CURRENT.Cells(i, 22).Value
        Z = CURRENT.Cells(i, 26).Value
        AA = CURRENT.Cells(i, 27).Value
        AB = CURRENT.Cells(i, 28).Value
        
        
        
    
    
        LastArrayValue = Len(CStr(CURRENT.Cells(i, 34).Value))
        FirstArrayValue = 1
        
        For j = FirstArrayValue To LastArrayValue
        
        
            If Mid(CURRENT.Cells(i, 34).Value, j, 1) <> "," Then
                    'concatonates the stowage term
                    CONCAT = CONCAT + Mid(CURRENT.Cells(i, 34).Value, j, 1)
                    
            
            End If
            
            'MsgBox (Mid(CURRENT.Cells(i, 34).Value, j, 1))
            
            If Mid(CURRENT.Cells(i, 34).Value, j, 1) = "," Then
            
                'posts whatever has been collected so far into the Output cell
                OUTPUT_CURRENT.Cells(k, 1) = CStr(AC)
                OUTPUT_CURRENT.Cells(k, 2).Value = CONCAT
                OUTPUT_CURRENT.Cells(k, 3).Value = F
                OUTPUT_CURRENT.Cells(k, 4).Value = K2
                OUTPUT_CURRENT.Cells(k, 5).Value = n
                OUTPUT_CURRENT.Cells(k, 6).Value = U
                OUTPUT_CURRENT.Cells(k, 7).Value = V
                OUTPUT_CURRENT.Cells(k, 8).Value = Z
                OUTPUT_CURRENT.Cells(k, 9).Value = AA
                OUTPUT_CURRENT.Cells(k, 10).Value = AB
                CONCAT = ""
                k = k + 1
                
                
                
   'go through the nested layers
                If CURRENT.Cells(i, 19).Value = 0 Then
                
                TempArray = ""
                k = k - 1 '
                y = 11 'this is the column indicator for the output
                
                    If CURRENT.Cells(i + 1, 19).Value <> 0 And CURRENT.Cells(i + 1, 19).Value <> "" Then
                    
                    p = i + 1 'reassigning i so I don't lose it's place
                    
                
                        Do Until TempArray = "FALSE"
                        
                            
                        
                            If CURRENT.Cells(p, 19).Value <> 0 And CURRENT.Cells(p, 19).Value <> "" Then
                                
                                
                                
                                TempArray = "TRUE"
                                
                                'Since I'm on the nested lines info I can go ahead and pull the info and add it on the OUTPUT
                                
                                    If CURRENT.Cells(p, 19).Value = 1 Then
                                    
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 14).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 15).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 28).Value
                                    y = y + 1
                                    
                                    Else
                                
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 28).Value
                                    y = y + 1
                                
                                
                                    End If
                                
                                
                                
                                
                                
                                
                                'the end
                                
                                
                                p = p + 1
                                
                                
                            Else
                                TempArray = "FALSE"
                                Exit Do
                            
                            End If
                            
                        Loop
                        
                    End If
                    
                    k = k + 1
                    
                                   
                End If
                
                
                
                
                
                
                
                
                If k <> 1 Then
                
                    If OUTPUT_CURRENT.Cells(k - 1, 1).Value = "" Then
    
                        k = k - 1
                        
    
                    End If
                
                End If
                
                
                
                
                
               ' Selection.Copy
                'range(CURRENT.Cells(i + 1, 1), CURRENT.Cells(i + 1, 42)).Select
                'Selection.Insert Shift:=xlDown
                
                
                
                
                If CONCAT = "" Then
                
                    k = k
                
                End If
                
                If k <> 1 Then
                
                    If CONCAT <> "" And OUTPUT_CURRENT.Cells(k - 1).Value <> "" Then
                
                    
                        If OUTPUT_CURRENT.Cells(k - 1, 1).Value <> "" Then
    
                        k = k + 1
                        
    
                        End If
                    
                        
                    
                        CONCAT = ""
                                  
                    End If
                
                End If
            
            End If
        
        Next j
    
    End If
    
    
    If CONCAT <> "" Then
        OUTPUT_CURRENT.Cells(k, 1).Value = CStr(AC)
        OUTPUT_CURRENT.Cells(k, 2).Value = CONCAT
        OUTPUT_CURRENT.Cells(k, 3).Value = F
        OUTPUT_CURRENT.Cells(k, 4).Value = K2
        OUTPUT_CURRENT.Cells(k, 5).Value = n
        OUTPUT_CURRENT.Cells(k, 6).Value = U
        OUTPUT_CURRENT.Cells(k, 7).Value = V
        OUTPUT_CURRENT.Cells(k, 8).Value = Z
        OUTPUT_CURRENT.Cells(k, 9).Value = AA
        OUTPUT_CURRENT.Cells(k, 10).Value = AB
        CONCAT = ""
        k = k + 1
        
        
        
        
        
        
        
        
           'go through the nested layers
                If CURRENT.Cells(i, 19).Value = 0 Then
                
                TempArray = ""
                k = k - 1 'this may not be right might have to be fixed somewhere else!!!!!!!!!!!!!!!!!!!!!!
                y = 11 'this is the column indicator for the output
                
                    If CURRENT.Cells(i + 1, 19).Value <> 0 And CURRENT.Cells(i + 1, 19).Value <> "" Then
                    
                    p = i + 1 'reassigning i so I don't lose it's place
                    
                
                        Do Until TempArray = "FALSE"
                        
                            
                        
                            If CURRENT.Cells(p, 19).Value <> 0 And CURRENT.Cells(p, 19).Value <> "" Then
                                
                                
                                
                                TempArray = "TRUE"
                                
                                'Since I'm on the nested lines info I can go ahead and pull the info and add it on the OUTPUT
                                
                                
                                    If CURRENT.Cells(p, 19).Value = 1 Then
                                    
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 14).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 15).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 28).Value
                                    y = y + 1
                                    
                                    Else
                                
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_CURRENT.Cells(k, y).Value = CURRENT.Cells(p, 28).Value
                                    y = y + 1
                                
                                
                                    End If
                                
                                
                                
                                
                                
                                
                                
                                'the end
                                
                                
                                p = p + 1
                                
                                
                            Else
                                TempArray = "FALSE"
                                Exit Do
                            
                            End If
                            
                        Loop
                        
                    End If
                    
                    k = k + 1
                    
                                   
                End If
        
        
        
        
        
        
        
        
        
    
        If k <> 1 Then
    
            If OUTPUT_CURRENT.Cells(k, 1).Value <> "" Then
    
                k = k + 1
                
    
            End If
        
        End If
        
    End If
    
    If CONCAT = "" Then
                
        k = k
                
    End If
                
    If k <> 1 Then
                
    If CONCAT <> "" And OUTPUT_CURRENT.Cells(k - 1, 1).Value <> "" Then
                
        
            If OUTPUT_CURRENT.Cells(k - 1, 1).Value <> "" Then
    
                k = k + 1
                
    
            End If
        
        
            CONCAT = ""
        
    
            End If
    
    End If
    

Next i

























'stowage list from future



FUTURE.Select
FUTURE.Cells(1, 1).Select
toprow2 = FUTURE.Cells(1, 1).Row
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row

OUTPUT_FUTURE.Select

k = 1
p = 1

For i = 2 To endrow1

    
    If CURRENT.Cells(i, 34).Value <> "" Then
    
        AC = CStr(FUTURE.Cells(i, 1).Value)
        F = FUTURE.Cells(i, 6).Value
        K2 = FUTURE.Cells(i, 11).Value
        n = FUTURE.Cells(i, 14).Value
        U = FUTURE.Cells(i, 21).Value
        V = FUTURE.Cells(i, 22).Value
        Z = FUTURE.Cells(i, 26).Value
        AA = FUTURE.Cells(i, 27).Value
        AB = FUTURE.Cells(i, 28).Value
        
        
        
    
    
        LastArrayValue = Len(CStr(FUTURE.Cells(i, 34).Value))
        FirstArrayValue = 1
        
        For j = FirstArrayValue To LastArrayValue
        
        
            If Mid(FUTURE.Cells(i, 34).Value, j, 1) <> "," Then
                    'concatonates the stowage term
                    CONCAT = CONCAT + Mid(FUTURE.Cells(i, 34).Value, j, 1)
                    
            
            End If
            
            'MsgBox (Mid(CURRENT.Cells(i, 34).Value, j, 1))
            
            If Mid(FUTURE.Cells(i, 34).Value, j, 1) = "," Then
            
                'posts whatever has been collected so far into the Output cell
                OUTPUT_FUTURE.Cells(k, 1) = CStr(AC)
                OUTPUT_FUTURE.Cells(k, 2).Value = CONCAT
                OUTPUT_FUTURE.Cells(k, 3).Value = F
                OUTPUT_FUTURE.Cells(k, 4).Value = K2
                OUTPUT_FUTURE.Cells(k, 5).Value = n
                OUTPUT_FUTURE.Cells(k, 6).Value = U
                OUTPUT_FUTURE.Cells(k, 7).Value = V
                OUTPUT_FUTURE.Cells(k, 8).Value = Z
                OUTPUT_FUTURE.Cells(k, 9).Value = AA
                OUTPUT_FUTURE.Cells(k, 10).Value = AB
                CONCAT = ""
                k = k + 1
                
                
                
   'go through the nested layers
                If FUTURE.Cells(i, 19).Value = 0 Then
                
                TempArray = ""
                k = k - 1 'this may not be right might have to be fixed somewhere else!!!!!!!!!!!!!!!!!!!!!!
                y = 11 'this is the column indicator for the output
                
                    If FUTURE.Cells(i + 1, 19).Value <> 0 And FUTURE.Cells(i + 1, 19).Value <> "" Then
                    
                    p = i + 1 'reassigning i so I don't lose it's place
                    
                
                        Do Until TempArray = "FALSE"
                        
                            
                        
                            If FUTURE.Cells(p, 19).Value <> 0 And FUTURE.Cells(p, 19).Value <> "" Then
                                
                                
                                
                                TempArray = "TRUE"
                                
                                'Since I'm on the nested lines info I can go ahead and pull the info and add it on the OUTPUT
                                
                                    If FUTURE.Cells(p, 19).Value = 1 Then
                                    
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 14).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 15).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 28).Value
                                    y = y + 1
                                    
                                    Else
                                
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 28).Value
                                    y = y + 1
                                
                                
                                    End If
                                
                                
                                
                                
                                
                                
                                'the end
                                
                                
                                p = p + 1
                                
                                
                            Else
                                TempArray = "FALSE"
                                Exit Do
                            
                            End If
                            
                        Loop
                        
                    End If
                    
                    k = k + 1
                    
                                   
                End If
                
                
                
                
                
                
                
                
                If k <> 1 Then
                
                    If OUTPUT_FUTURE.Cells(k - 1, 1).Value = "" Then
    
                        k = k - 1
                        
    
                    End If
                
                End If
                
                
                
                
                
               ' Selection.Copy
                'range(CURRENT.Cells(i + 1, 1), CURRENT.Cells(i + 1, 42)).Select
                'Selection.Insert Shift:=xlDown
                
                
                
                
                If CONCAT = "" Then
                
                    k = k
                
                End If
                
                If k <> 1 Then
                
                    If CONCAT <> "" And OUTPUT_FUTURE.Cells(k - 1).Value <> "" Then
                
                    
                        If OUTPUT_FUTURE.Cells(k - 1, 1).Value <> "" Then
    
                        k = k + 1
                        
    
                        End If
                    
                        
                    
                        CONCAT = ""
                                  
                    End If
                
                End If
            
            End If
        
        Next j
    
    End If
    
    
    If CONCAT <> "" Then
        OUTPUT_FUTURE.Cells(k, 1).Value = CStr(AC)
        OUTPUT_FUTURE.Cells(k, 2).Value = CONCAT
        OUTPUT_FUTURE.Cells(k, 3).Value = F
        OUTPUT_FUTURE.Cells(k, 4).Value = K2
        OUTPUT_FUTURE.Cells(k, 5).Value = n
        OUTPUT_FUTURE.Cells(k, 6).Value = U
        OUTPUT_FUTURE.Cells(k, 7).Value = V
        OUTPUT_FUTURE.Cells(k, 8).Value = Z
        OUTPUT_FUTURE.Cells(k, 9).Value = AA
        OUTPUT_FUTURE.Cells(k, 10).Value = AB
        CONCAT = ""
        k = k + 1
        
        
        
        'I need to RECODE the position and AC check to see if any of the positions are missing. XOXO
        
        
        
        
           'go through the nested layers
                If FUTURE.Cells(i, 19).Value = 0 Then
                
                TempArray = ""
                k = k - 1 'this may not be right might have to be fixed somewhere else!!!!!!!!!!!!!!!!!!!!!!
                y = 11 'this is the column indicator for the output
                
                    If FUTURE.Cells(i + 1, 19).Value <> 0 And FUTURE.Cells(i + 1, 19).Value <> "" Then
                    
                    p = i + 1 'reassigning i so I don't lose it's place
                    
                
                        Do Until TempArray = "FALSE"
                        
                            
                        
                            If FUTURE.Cells(p, 19).Value <> 0 And FUTURE.Cells(p, 19).Value <> "" Then
                                
                                
                                
                                TempArray = "TRUE"
                                
                                'Since I'm on the nested lines info I can go ahead and pull the info and add it on the OUTPUT
                                
                                
                                    If FUTURE.Cells(p, 19).Value = 1 Then
                                    
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 14).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 15).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 28).Value
                                    y = y + 1
                                    
                                    Else
                                
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 11).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 21).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 22).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 26).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 27).Value
                                    y = y + 1
                                    OUTPUT_FUTURE.Cells(k, y).Value = FUTURE.Cells(p, 28).Value
                                    y = y + 1
                                
                                
                                    End If
                                
                                
                                
                                
                                
                                
                                
                                'the end
                                
                                
                                p = p + 1
                                
                                
                            Else
                                TempArray = "FALSE"
                                Exit Do
                            
                            End If
                            
                        Loop
                        
                    End If
                    
                    k = k + 1
                    
                                   
                End If
        
        
        
        
        
        
        
        
        
    
        If k <> 1 Then
    
            If OUTPUT_FUTURE.Cells(k, 1).Value <> "" Then
    
                k = k + 1
                
    
            End If
        
        End If
        
    End If
    
    If CONCAT = "" Then
                
        k = k
                
    End If
                
    If k <> 1 Then
                
    If CONCAT <> "" And OUTPUT_FUTURE.Cells(k - 1, 1).Value <> "" Then
                
        
            If OUTPUT_FUTURE.Cells(k - 1, 1).Value <> "" Then
    
                k = k + 1
                
    
            End If
        
        
            CONCAT = ""
        
    
            End If
    
    End If
    

Next i






'check stowage positions here














Dim MATCHER As String



k = 1

'compare output_current to output_future


OUTPUT_CURRENT.Select
OUTPUT_CURRENT.Cells(1, 1).Select
toprow1 = OUTPUT_CURRENT.Cells(1, 1).Row
Selection.End(xlDown).Select
endrow1 = ActiveCell.Row


OUTPUT_FUTURE.Select
OUTPUT_FUTURE.Cells(1, 1).Select
toprow2 = OUTPUT_FUTURE.Cells(1, 3).Row
Selection.End(xlDown).Select
endrow2 = ActiveCell.Row

For i = toprow1 To endrow1

    MATCHER = "FALSE"

    For j = toprow2 To endrow2
    
        If OUTPUT_CURRENT.Cells(i, 1).Value = OUTPUT_FUTURE.Cells(j, 1).Value And OUTPUT_CURRENT.Cells(i, 2).Value = OUTPUT_FUTURE.Cells(j, 2).Value Then
            
            For p = 3 To 400 'this is where we compare all the nested stuff
            
                If OUTPUT_CURRENT.Cells(i, p).Value <> OUTPUT_FUTURE.Cells(j, p).Value Then
                
                    DISCREPANCIES.Cells(k, 1).Value = OUTPUT_CURRENT.Cells(i, 1).Value
                    DISCREPANCIES.Cells(k, 2).Value = OUTPUT_CURRENT.Cells(i, 2).Value
                    DISCREPANCIES.Cells(k, 3).Value = "!!!CURRENT value is " + CStr(OUTPUT_CURRENT.Cells(i, p).Value) + " !!!FUTURE value is " + CStr(OUTPUT_FUTURE.Cells(j, p).Value)
                    k = k + 1
                
                End If
            
            Next p
            
            
            MATCHER = "TRUE"
            Exit For
            
        End If
             
    Next j
    
    If MATCHER = "FALSE" Then
        
        
        k = k + 1
    
    End If

Next i




If DISCREPANCIES.Cells(1, 1).Value = "" Then

    DISCREPANCIES.Cells(1, 1).Value = "NO CAHNGES FOUND!!!"

End If













End Sub





Sub changes_delete_blanks()

Dim Sheet2 As Worksheet
Set Sheet2 = Worksheets("Sheet2")

Dim i As Long
Dim k As Long

k = 1

For i = 1 To 2000

    If Sheet2.Cells(i, 1).Value = "" Then
        Sheet2.Cells(i, 1).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Delete Shift:=xlUp
        i = i - 1
        k = k + 1
    End If
    
    
    
    If k > 100 Then
        Exit For
    End If

Next i






End Sub





Sub RUN_BILLCODE()

Dim PASTE_BILLCODE_RUN_BILLCODE As Worksheet
Set PASTE_BILLCODE_RUN_BILLCODE = Worksheets("PASTE_BILLCODE_RUN_BILLCODE")

Dim i As Long
Dim j As Long

    PASTE_BILLCODE_RUN_BILLCODE.Cells(1, 1).Select
    Selection.End(xlDown).Select
    endrow = ActiveCell.Row
    
    For i = 1 To endrow
    
        On Error GoTo ErrHandler:
    
        
        
            'below appends value of tail numbers into an array
            LastArrayValue = Len(CStr(PASTE_BILLCODE_RUN_BILLCODE.Cells(i, 1).Value))
            FirstArrayValue = 1
            
        
            For j = FirstArrayValue To LastArrayValue
                
                
                If Mid(PASTE_BILLCODE_RUN_BILLCODE.Cells(i, 1).Value, j, 1) <> "-" Then
                    
                    
                    
                ElseIf Mid(PASTE_BILLCODE_RUN_BILLCODE.Cells(i, 1).Value, j, 1) = "-" Then
                    PASTE_BILLCODE_RUN_BILLCODE.Cells(i, 2).Value = Mid(PASTE_BILLCODE_RUN_BILLCODE.Cells(i, 1).Value, 1, j - 1)
                    
                    GoTo nextthing:
                    
                
                 End If
                
            
                

            Next j
        
    
ErrHandler:
nextthing:
   
                    
        
    
    Next i
    
    
 
    
    
    
    
    
    
    
    
    
    
    
    
    


End Sub




Sub InsertCommas()

Dim M400 As Worksheet
Set M400 = Worksheets("M400")

Dim i As Long
Dim j As Long
Dim k As Long

M400.Cells(1, 1).Select
topRow = M400.Cells(1, 1).Row
Selection.End(xlDown).Select
endrow = ActiveCell.Row




For i = 2 To 70 'NUMBER OF TIMES REPEATED
    
    
    M400.Cells(1, i).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert Shift:=xlToRight
    M400.Cells(1, i).Select
    
    For j = topRow To endrow
    
        M400.Cells(j, i).Value = "$"
        
    
    Next j
    
    
    i = i + 1
    
    


Next i




End Sub


Sub COLUMN_MERGE()

'DO THIS AFTER TO APPLY DELIMITER COLUMN SEPERATOR
'copy paste transpose
'NOTE NAME YOUR SHEET COLUMN_MERGE

'WHEN IT FINISHES COPY THE ORIGINAL CONTAINER_NAME AND STOWAGE COLUMNS INTO COLUMN_MERGE COLUMN D1
' THE LIST OF TRANSPOSED STOWAGES SHOULD BE ON THE FAR LEFT AND THE ORIGINAL STOWAGES AND DESCRIPTIONS SHOULD BE ON RIGHT
'THEN CONTINUE LOOP

Dim COLUMN_MERGE As Worksheet
Set COLUMN_MERGE = Worksheets("COLUMN_MERGE")

Dim COLUMN_MERGE2 As Worksheet
Set COLUMN_MERGE2 = Worksheets("COLUMN_MERGE2")

Dim i As Long
Dim j As Long
Dim k As Long

COLUMN_MERGE.Select
Range("a1").Select
Selection.End(xlDown).Select
k = Selection.Row

k = k + 1

Dim FIRSTCOLUMN As Long
Dim LASTCOLUMN As Long

FIRSTCOLUMN = 2
LASTCOLUMN = 59

Dim FIRSTROW As Long
Dim LASTROW As Long

FIRSTROW = 1
LASTROW = 12

 
For i = FIRSTCOLUMN To LASTCOLUMN
    For j = FIRSTROW To LASTROW
        
        If COLUMN_MERGE.Cells(j, i).Value <> "" Then
        
            COLUMN_MERGE.Cells(k, 1).Value = COLUMN_MERGE.Cells(j, i).Value
            k = k + 1
            
        ElseIf COLUMN_MERGE.Cells(j, i).Value = "" Then
        
            Exit For
         
        End If
        
        COLUMN_MERGE.Cells(j, i).Value = ""
        
    Next j
Next i

'WHEN IT FINISHES COPY THE ORIGINAL CONTAINER_NAME AND STOWAGE COLUMNS INTO COLUMN_MERGE COLUMN D1
'THEN CONTINUE LOOP

FIRSTROW = 1  'THIS STOPPER NEEDS TO STAY HERE!!!!!!!!!
LASTROW = 1000

k = 1

Dim COUNTER As Integer
Dim PASTER As Integer


For j = FIRSTROW To LASTROW

COUNTER = 1
        
        If COLUMN_MERGE.Cells(j, 5).Value <> "" And COLUMN_MERGE.Cells(j, 5).Value <> "STOWAGE_LIST" Then
        
        
            For i = 1 To Len(COLUMN_MERGE.Cells(j, 5).Value)
            
    
                If Mid(COLUMN_MERGE.Cells(j, 5).Value, i, 1) = "," Then
                    COUNTER = COUNTER + 1
                End If
                
            Next i
            
            
            
            For PASTER = 1 To COUNTER
            
                COLUMN_MERGE.Cells(k, 2).Value = COLUMN_MERGE.Cells(j, 4).Value
                k = k + 1
            
            Next PASTER
            
        End If
        
Next j



End Sub


Sub replaceBlanksOracle()

Dim IFX As Worksheet
Set IFX = Worksheets("IFX")

Dim i As Long
Dim j As Long
Dim k As Long



startrow = 1

Range("a1").Select
Selection.End(xlToRight).Select

lastcol = 12

Selection.End(xlDown).Select

LASTROW = Selection.Row


For i = startrow To LASTROW

    For j = 1 To lastcol
    
        If IFX.Cells(i, j).Value = "" Then
            IFX.Cells(i, j).Value = "X"
        End If
    
    Next j

Next i

End Sub



Sub replacenumberwithairportcode()

Dim Sheet1 As Worksheet
Set Sheet1 = Worksheets("Sheet1")

Dim i As Long
Dim j As Long

startrow1 = Sheet1.Cells(1, 1).Row
Sheet1.Select
Range("b1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row


For i = startrow1 To lastrow1

    If Sheet1.Cells(i, 5).Value = 1616 Then
        Sheet1.Cells(i, 5).Value = "ANC"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1060 Then
        Sheet1.Cells(i, 5).Value = "BWI"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1468 Then
        Sheet1.Cells(i, 5).Value = "CLT"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1026 Then
        Sheet1.Cells(i, 5).Value = "DCA"
        
    ElseIf Sheet1.Cells(i, 5).Value = 235 Then
        Sheet1.Cells(i, 5).Value = "DEN"
        
    ElseIf Sheet1.Cells(i, 5).Value = 195 Then
        Sheet1.Cells(i, 5).Value = "DFW"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1399 Then
        Sheet1.Cells(i, 5).Value = "DTW"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1059 Then
        Sheet1.Cells(i, 5).Value = "IAD"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1316 Then
        Sheet1.Cells(i, 5).Value = "IAD"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1371 Then
        Sheet1.Cells(i, 5).Value = "JFK"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1479 Then
        Sheet1.Cells(i, 5).Value = "MCO"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1366 Then
        Sheet1.Cells(i, 5).Value = "MIA"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1390 Then
        Sheet1.Cells(i, 5).Value = "MSP"
        
    ElseIf Sheet1.Cells(i, 5).Value = 710 Then
        Sheet1.Cells(i, 5).Value = "PDX"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1376 Then
        Sheet1.Cells(i, 5).Value = "PHL"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1692 Then
        Sheet1.Cells(i, 5).Value = "PHX"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1375 Then
        Sheet1.Cells(i, 5).Value = "PIT"
        
    ElseIf Sheet1.Cells(i, 5).Value = 460 Then
        Sheet1.Cells(i, 5).Value = "RDU"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1481 Then
        Sheet1.Cells(i, 5).Value = "RSW"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1393 Then
        Sheet1.Cells(i, 5).Value = "SJC"
    
     ElseIf Sheet1.Cells(i, 5).Value = 1307 Then
        Sheet1.Cells(i, 5).Value = "SLC"
        
    ElseIf Sheet1.Cells(i, 5).Value = 1381 Then
        Sheet1.Cells(i, 5).Value = "SNA"
        
    ElseIf Sheet1.Cells(i, 5).Value = 260 Then
        Sheet1.Cells(i, 5).Value = "AUS"
        
    End If


Next i

End Sub

Sub tailNumberAudit2()



Dim IFX_FLEET_LIST As Worksheet
Set IFX_FLEET_LIST = Worksheets("IFX_FLEET_LIST")

Dim CBASE_TAIL_NUMBER As Worksheet
Set CBASE_TAIL_NUMBER = Worksheets("CBASE_TAIL_NUMBER")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Integer
Dim w As Long
Dim AC As String
Dim x As Integer
Dim TAIL As String
Dim y As Integer
Dim AC_VAL As Variant
Dim AAC_VAL As Variant
Dim Count_space As Integer
Dim FINISH As Variant
Dim aac As Variant


Dim str As Long
Dim FirstArrayValue As Long
Dim LastArrayValue As Long
Dim TAIL_VALUE As Variant

w = 1

AC = "31J,3HF,32K/3KR,32M/3MR,321,333,332,33X,359,717,738,73A,73W,739,73E,75G,75S,75C,75D,75H,75Y,76Z,76P,76T,76D,76L,7HD,777,77B,M88,M90,CJ7,CAJ,CPJ,CM9,CM7,E70,E75,RJ7,RJ9,ES4,ES5,RJW,RJ8 F,CG7,CG7"

aac = Split(AC, ",")



COUNTER = 0



    For j = 1 To Len(AC)
        If Mid(AC, j, 1) = "," Then
   
            COUNTER = COUNTER + 1

        End If
        
    Next j


AC_LEN = COUNTER




Columns("A:A").Select
    Selection.Replace What:=" - ", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" -", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="- ", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False











startrow1 = IFX_FLEET_LIST.Cells(1, 1).Row
IFX_FLEET_LIST.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row
Selection.End(xlUp).Select



For i = startrow1 To lastrow1

    For j = 0 To AC_LEN

        If IFX_FLEET_LIST.Cells(i, 1).Value = aac(j) Then
        
            If aac(j) = "321" Or aac(j) = "CPJ" Or aac(j) = "E70" Or aac(j) = "RJ8 F" Then
                
                AC_VAL = aac(j)
                i = i + 1
                
                
                
            Else
            
                i = i + 2
                
                
                AC_VAL = aac(j)
                TAIL_VALUE = IFX_FLEET_LIST.Cells(i, 1).Value
                'MsgBox (TAIL_VALUE)
                
                For k = 1 To Len(TAIL_VALUE)
                
                     'MsgBox (Mid(TAIL_VALUE, k, 1))
                    TAIL = ""
                    'MsgBox (Len(TAIL_VALUE))
                    'MsgBox (k)
                    'MsgBox (Mid(TAIL_VALUE, k, 1))
                    
                
           
                Do While IsNumeric(Mid(TAIL_VALUE, k, 1))
                    'MsgBox (Mid(TAIL_VALUE, k, 1))
                    TAIL = CStr(TAIL) + CStr(Mid(TAIL_VALUE, k, 1))
                   'MsgBox (TAIL)
                   k = k + 1
                    
                Loop
                
                    OUTPUT.Cells(w, 1).Value = AC_VAL
                    OUTPUT.Cells(w, 2).Value = TAIL
                    w = w + 1
                    
                    If OUTPUT.Cells(w - 1, 2).Value = "" And OUTPUT.Cells(w - 1, 1).Value <> "" Then
                                
                        OUTPUT.Cells(w - 1, 1).Value = ""
                        w = w - 1
                        
                                      
                    End If
                    
                    
                
                If Mid(TAIL_VALUE, k, 1) = "," Or Mid(TAIL_VALUE, k, 1) = " " Then
                
                    k = k + 1
                    
                
                                      
                End If
                
                
                
                
                
                
                
                
                
             
                
                    
                    
                    
                    
                    
                    
                    
                If Mid(TAIL_VALUE, k, 1) = "-" Then
                                
                       START = OUTPUT.Cells(w - 1, 2).Value
                       
                       k = k + 1
                       
                       Do While IsNumeric(Mid(TAIL_VALUE, k, 1))
                       
                            FINISH = CStr(FINISH) + CStr(Mid(TAIL_VALUE, k, 1))
                            k = k + 1
                       
                       Loop
                       
                       'MsgBox (START1)
                       'MsgBox (FINISH)
                       
                       TAIL = CInt(START) + 1
                       START1 = CInt(START)
                       
                       For y = START1 + 1 To FINISH
                            
                            OUTPUT.Cells(w, 1).Value = AC_VAL
                            OUTPUT.Cells(w, 2).Value = TAIL
                            w = w + 1
                            TAIL = TAIL + 1
                            
                            
                            
                            
                       
                       Next y
                       
                       FINISH = 0
                       
                       k = k - 1
                
                End If
                    
                
              'MsgBox (Mid(TAIL_VALUE, k, 1))
                
             Next k
                
                
                
                
                
                
                
                
                
            
            
            End If
        
            
            
            
            'break apart tail numbers by comma here
            
            
                
                
                
                
                
                
                
                
                
                
                
                
            
            
        End If
        'MsgBox (aac(j))
        'MsgBox (IFX_FLEET_LIST.Cells(i, 1).Value)
        
        

    Next j
    

Next i

End Sub





Sub tailNumberAudit()

Dim IFX_FLEET_LIST As Worksheet
Set IFX_FLEET_LIST = Worksheets("IFX_FLEET_LIST")

Dim CBASE_TAIL_NUMBER As Worksheet
Set CBASE_TAIL_NUMBER = Worksheets("CBASE_TAIL_NUMBER")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Integer
Dim w As Long
Dim AC As String
Dim x As Integer
Dim TAIL As Long
Dim y As Integer
Dim AC_VAL As Variant
Dim AAC_VAL As Variant
Dim Count_space As Integer




Dim str As Long
Dim FirstArrayValue As Long
Dim LastArrayValue As Long

i = 1
j = 1
k = 1


'index all aircraft

AC = "31J,321,32K,32K/3KR,32M/3MR,332,333,359,717,738,739/E,73A,73W,75C,75DH,75G,75S,75Y,76D,76L,76PQ,76T,76Z,777,77B,7HD,CG7,CG9,CJ7,CJ9,CM7,CM9,CPJ,E70,E75,EC5,ES4,ES5,M88,M90,RJ7,RJ8,RJ9,RJS"

aac = Split(AC, ",")


COUNTER = 0



    For j = 1 To Len(AC)
        If Mid(AC, j, 1) = "," Then
   
            COUNTER = COUNTER + 1

        End If
        
    Next j


AC_LEN = COUNTER

Dim FINISH As Variant




startrow1 = IFX_FLEET_LIST.Cells(1, 1).Row
IFX_FLEET_LIST.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row
Selection.End(xlUp).Select

TAIL = 1

For i = 4 To lastrow1


        
    For j = 0 To AC_LEN
    
 
       
       If IFX_FLEET_LIST.Cells(i, 1).Value = aac(j) Then
       
            AC_VAL = aac(j)
            i = i + 1
            
            For x = 1 To 4
            
                If Mid(IFX_FLEET_LIST.Cells(i, 1).Value, 1) <> "," Then
                
                Else
                    i = i + 1
                    
                End If
                
            
            Next x
            
                For k = 1 To Len(IFX_FLEET_LIST.Cells(i, 1).Value)
                
                
            
                    If Mid(IFX_FLEET_LIST.Cells(i, 1).Value, k, 1) = "," Then
                
                        'break down all the tail numbers here
                        
                        AAC_VAL = ""
                        
                        For y = 1 To Len(IFX_FLEET_LIST.Cells(i, 1).Value)
                        
                             If IsNumeric(Mid(IFX_FLEET_LIST.Cells(i, 1).Value, y, 1)) Then
                            
                            
                                AAC_VAL = AAC_VAL + Mid(IFX_FLEET_LIST.Cells(i, 1).Value, y, 1)
                                
                            Else
                            
                                OUTPUT.Cells(TAIL, 1).Value = AC_VAL
                                OUTPUT.Cells(TAIL, 2).Value = AAC_VAL
                                AAC_VAL = ""
                                TAIL = TAIL + 1
                                
                                If OUTPUT.Cells(TAIL - 1, 2).Value = "" And OUTPUT.Cells(TAIL - 1, 1).Value <> "" Then
                                
                                    OUTPUT.Cells(TAIL - 1, 1).Value = ""
                                    TAIL = TAIL - 1
                                      
                                End If
                                
                            End If
                                       
  
                                
                            If Mid(IFX_FLEET_LIST.Cells(i, 1).Value, y, 1) = "-" Then
                                

                                START = OUTPUT.Cells(TAIL - 1, 2).Value
                                'MsgBox (START)
                               
                                
                                n = y + 1
                                
                                Do While IsNumeric(Count_space)
                                
                                    If IsNumeric(Mid(IFX_FLEET_LIST.Cells(i, 1).Value, n, 1)) Then
                                        Count_space = Mid(IFX_FLEET_LIST.Cells(i, 1).Value, n, 1)
                                        FINISH = CStr(FINISH) + CStr(Count_space)
                                        n = n + 1
                                        
                                    ElseIf IsNumeric(Mid(IFX_FLEET_LIST.Cells(i, 1).Value, n + 1, 1)) Then

                                        n = n + 1
                                        
                                    ElseIf Mid(IFX_FLEET_LIST.Cells(i, 1).Value, n, 1) = "," Or Mid(IFX_FLEET_LIST.Cells(i, 1).Value, n, 1) = " " Then
                                    
                                        'count out the tail numbers here
                                            
                                        
                                        For w = START + 1 To FINISH
                                        
                                        
                                        
                                            OUTPUT.Cells(TAIL, 1).Value = AC_VAL
                                            OUTPUT.Cells(TAIL, 2).Value = w
                                            TAIL = TAIL + 1
                                            FINISH = ""
                                            
                                        
                                        Next w
                                        
                                        Exit Do
                                        
                                    Else
                                    
                                      Exit Do
                                      
                                        
                                    End If
                                
                                Loop
                                
                  
                                
                            End If
                                

                        Next y
                        
   
                    End If
                    
                    
                        OUTPUT.Cells(TAIL, 1).Value = AC_VAL
                        OUTPUT.Cells(TAIL, 2).Value = AAC_VAL
                        AAC_VAL = ""
                        TAIL = TAIL + 1
                        
                        If OUTPUT.Cells(TAIL - 1, 2).Value = "" And OUTPUT.Cells(TAIL - 1, 1).Value <> "" Then
                                
                            OUTPUT.Cells(TAIL - 1, 1).Value = ""
                            TAIL = TAIL - 1
                                      
                        End If
            
                Next k
                
                
            
            
            
           
        End If

    
    Next j


Next i





End Sub




Sub COMPARE_CBASE_GP4()

Dim DROP_CBASE_DIAGRAM_HERE As Worksheet
Set DROP_CBASE_DIAGRAM_HERE = Worksheets("DROP_CBASE_DIAGRAM_HERE")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim i As Long
Dim j As Long
Dim q As Integer
Dim x As String
Dim k As String

i = 1
j = 1

'specific loading macro needs to be run before this one


For i = 1 To 500

    DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value = DROP_CBASE_DIAGRAM_HERE.Cells(i, 4).Value
    

Next i

For i = 1 To 500

    If trim(DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value) = "Number" Then
        DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value = ""
    End If

Next i

'get rid of blanks

For i = 1 To 500

    If DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value = "" Then
        For j = i To 500
            DROP_CBASE_DIAGRAM_HERE.Cells(j, 12).Value = DROP_CBASE_DIAGRAM_HERE.Cells(j + 1, 12).Value
        Next j
    End If

Next i

For i = 1 To 500

    If DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value = "" Then
        For j = i To 500
            DROP_CBASE_DIAGRAM_HERE.Cells(j, 12).Value = DROP_CBASE_DIAGRAM_HERE.Cells(j + 1, 12).Value
        Next j
    End If

Next i

'get rid of duplicates

    Columns("L:L").Select
    ActiveSheet.Range("$L$1:$L$500").RemoveDuplicates Columns:=1, HEADER:=xlNo
    
    'take stowages and add them to a variable
    
startrow1 = DROP_CBASE_DIAGRAM_HERE.Cells(1, 1).Row
DROP_CBASE_DIAGRAM_HERE.Select
Range("l1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row
    
    For i = startrow1 To lastrow1
    
        x = ""
    
        For j = 1 To 500
        
            If DROP_CBASE_DIAGRAM_HERE.Cells(i, 12).Value = DROP_CBASE_DIAGRAM_HERE.Cells(j, 4).Value Then
                
                x = x + "/" + DROP_CBASE_DIAGRAM_HERE.Cells(j, 1).Value
                DROP_CBASE_DIAGRAM_HERE.Cells(i, 13).Value = x
                
            End If
        
        Next j
        
        
    
    Next i
    
    
startrow1 = OUTPUT.Cells(1, 1).Row
OUTPUT.Select
Range("b1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row


startrow2 = DROP_CBASE_DIAGRAM_HERE.Cells(1, 1).Row
DROP_CBASE_DIAGRAM_HERE.Select
Range("l1").Select
Selection.End(xlDown).Select
lastrow2 = Selection.Row
    
    
 For i = startrow1 To lastrow1
    For j = startrow2 To lastrow2
    
        If OUTPUT.Cells(i, 3).Value = DROP_CBASE_DIAGRAM_HERE.Cells(j, 12).Value Then
            OUTPUT.Cells(i, 4).Value = DROP_CBASE_DIAGRAM_HERE.Cells(j, 13).Value
        End If
    
    Next j
 Next i
    

End Sub




Sub SpecificLoading()

Dim DROP_FLIGHT_SPEC_GP4_HERE As Worksheet
Set DROP_FLIGHT_SPEC_GP4_HERE = Worksheets("DROP_FLIGHT_SPEC_GP4_HERE")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim MATCHED_CODES As Worksheet
Set MATCHED_CODES = Worksheets("MATCHED_CODES")

Dim i As Long
Dim j As Long
Dim q As Integer
Dim x As String
Dim k As String


i = 1
j = 1



'make multiple container array here

Dim MULTCON(0 To 4) As String

MULTCON(0) = "PCHD"
MULTCON(1) = "PD6D"
MULTCON(2) = "FD6D"
MULTCON(3) = "BD8S"
MULTCON(4) = "PCH3"


startrow1 = DROP_FLIGHT_SPEC_GP4_HERE.Cells(1, 1).Row
DROP_FLIGHT_SPEC_GP4_HERE.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row

'look through col m
'if 0 then hold value of AK - 37 in a variable

'if 1 or 2 list value of variable then value from column N

For i = startrow1 To lastrow1
          
                'multiple contianer array check here
                
                For q = 0 To 4
                
                    If Mid(DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 9).Value, 1, 4) = MULTCON(q) Then '1
                    'here remember end if
                        x = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 37).Value
                        k = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 9).Value
                        
                        Do While DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 9).Value = k
                        
                                       
                            If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 14).Value <> "" Then
                            
                                If OUTPUT.Cells(j, 1).Value <> "" Then
                                
                                    j = j + 1
                                    
                                Else
                                
                                    OUTPUT.Cells(j, 1).Value = x
                                    OUTPUT.Cells(j, 2).Value = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 14).Value
                                    j = j + 1
                                    i = i + 1
                                
                                End If
                                
                            Else
                            
                                 i = i + 1
                                 
                                 If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i + 1, 22).Value = 0 Then
                        
                                    Exit Do
                        
                                End If
                                                            
                            End If
                        
                        Loop
                        
                    End If '1
                        
                Next q
                
                
                If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 22).Value = 0 Then
                    
                    x = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 37).Value
                    
                        If OUTPUT.Cells(j, 1).Value <> "" Then
                                
                            j = j + 1
                                    
                        Else
                                
                            OUTPUT.Cells(j, 1).Value = x
                            OUTPUT.Cells(j, 2).Value = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 9).Value
                            j = j + 1
                            i = i + 1
                                
                        End If
                        
                        If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 22).Value = 1 Then
                        
                            If OUTPUT.Cells(j, 1).Value <> "" Then
                                
                                j = j + 1
                                    
                            Else
                                
                                OUTPUT.Cells(j, 1).Value = x
                                OUTPUT.Cells(j, 2).Value = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 14).Value
                                j = j + 1
                                i = i + 1
                                
                            End If
                        
                        End If
                        
                        If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 22).Value = 2 Then
                        
                            If OUTPUT.Cells(j, 1).Value <> "" Then
                                
                                j = j + 1
                                    
                            Else
                                
                                OUTPUT.Cells(j, 1).Value = x
                                OUTPUT.Cells(j, 2).Value = DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 14).Value
                                j = j + 1
                                i = i + 1
                                
                            End If
                        
                        End If
                        
                        If DROP_FLIGHT_SPEC_GP4_HERE.Cells(i, 22).Value = 0 Then
                        
                            i = i - 1
                        
                        End If
                    
                
                End If


Next i

startrow1 = MATCHED_CODES.Cells(1, 1).Row
MATCHED_CODES.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row

startrow2 = OUTPUT.Cells(1, 1).Row
OUTPUT.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow2 = Selection.Row


For i = startrow2 To lastrow2

    For j = startrow1 To lastrow1
    
        If OUTPUT.Cells(i, 2).Value = MATCHED_CODES.Cells(j, 1).Value Then
        
            OUTPUT.Cells(i, 3).Value = MATCHED_CODES.Cells(j, 2).Value
        
        End If
    
    Next j


Next i




' after list match the list to list in Matched codes







End Sub


Sub CBASE_IFX_POSSIBLE_MATCH()

Dim DROP_FLIGHT_SPEC_GP4_HERE As Worksheet
Set DROP_FLIGHT_SPEC_GP4_HERE = Worksheets("DROP_FLIGHT_SPEC_GP4_HERE")

Dim OUTPUT As Worksheet
Set OUTPUT = Worksheets("OUTPUT")

Dim MATCHED_CODES As Worksheet
Set MATCHED_CODES = Worksheets("MATCHED_CODES")

Dim CBASE_FLIGHT_SPECIFIC As Worksheet
Set CBASE_FLIGHT_SPECIFIC = Worksheets("CBASE_FLIGHT_SPECIFIC")

Dim i As Long
Dim j As Long
Dim q As Integer
Dim wordCount As Integer
Dim x As String
Dim k As String
Dim position As Integer


i = 1
j = 1

startrow1 = OUTPUT.Cells(1, 1).Row 'j
OUTPUT.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow1 = Selection.Row

startrow2 = CBASE_FLIGHT_SPECIFIC.Cells(1, 1).Row
CBASE_FLIGHT_SPECIFIC.Select
Range("a1").Select
Selection.End(xlDown).Select
lastrow2 = Selection.Row

'Mid(Flightoverview.Cells(i, 2).Value, 3, Len(Flightoverview.Cells(i, 2).Value)



    For j = startrow1 To lastrow1 ' loop through Output
    
        x = Len(OUTPUT.Cells(j, 1).Value)
        
        For q = 1 To x
        
        k = ""
        
            If Mid(OUTPUT.Cells(j, 1).Value, q, 1) = " " Then
            
                wordCount = q + 3
                
                For position = wordCount To wordCount + 5
                
                If Mid(OUTPUT.Cells(j, 1).Value, position, 1) = "," Or Mid(OUTPUT.Cells(j, 1).Value, position, 1) = " " Then
                
                    Exit For
                
                End If
                
                k = k + Mid(OUTPUT.Cells(j, 1).Value, position, 1)
                
                
                
                            
                Next position
                
            
            End If
            
            If k <> "" Then
            
                Exit For
                
              
                
            End If
        
        
               
        Next q
        
        For i = startrow1 To lastrow1
        
        If k = CBASE_FLIGHT_SPECIFIC.Cells(i, 1).Value Then
        
        
            OUTPUT.Cells(j, 4).Value = CBASE_FLIGHT_SPECIFIC.Cells(i, 4).Value
            OUTPUT.Cells(j, 5).Value = CBASE_FLIGHT_SPECIFIC.Cells(i, 5).Value
             

        
        End If
        
       ' MsgBox (k)
        
        
        
        Next i
    
        
    
    Next j





End Sub





Sub match_pos_333()

Dim GP4_SPECIFIC As Worksheet
Set GP4_SPECIFIC = Worksheets("GP4_SPECIFIC")

Dim CBASE_SPECIFIC As Worksheet
Set CBASE_SPECIFIC = Worksheets("CBASE_SPECIFIC")

Dim GP4_pos(0 To 142) As String

GP4_pos(0) = "STOWAGE_LIST"
GP4_pos(1) = "A1 - A105"
GP4_pos(2) = " A1 - A107"
GP4_pos(3) = " A1 - A108"
GP4_pos(4) = " A1 - A109"
GP4_pos(5) = " A1 - A110"
GP4_pos(6) = " A1 - A111"
GP4_pos(7) = " A1 - A112"
GP4_pos(8) = "M1 - M106"
GP4_pos(9) = " M1 - M107"
GP4_pos(10) = " M1 - M109"
GP4_pos(11) = "M2 - M206F"
GP4_pos(12) = "M1 - M108"
GP4_pos(13) = "M1 - M103"
GP4_pos(14) = "A2 - A219F"
GP4_pos(15) = " A2 - A226F"
GP4_pos(16) = " A3 - A319F"
GP4_pos(17) = " A3 - A326F"
GP4_pos(18) = "M3 - M306F"
GP4_pos(19) = " M4 - M406F"
GP4_pos(20) = "M4 - M407F"
GP4_pos(21) = "F1 - F106F"
GP4_pos(22) = "M3 - M307F"
GP4_pos(23) = "A1 - A113"
GP4_pos(24) = "A1 - A116"
GP4_pos(25) = "M1 - M116"
GP4_pos(26) = "A2- - A228"
GP4_pos(27) = " A3- - A328"
GP4_pos(28) = "A2- - A229"
GP4_pos(29) = "A3- - A329"
GP4_pos(30) = "ZA - FD"
GP4_pos(31) = "M1 - M111"
GP4_pos(32) = "M2 - M206R"
GP4_pos(33) = "M1 - M110"
GP4_pos(34) = "M1 - M104"
GP4_pos(35) = "M2 - M204R"
GP4_pos(36) = "M2 - M202F"
GP4_pos(37) = "M2 - M205F"
GP4_pos(38) = "M1 - M105"
GP4_pos(39) = "M2 - M203R"
GP4_pos(40) = "A2 - A220F"
GP4_pos(41) = " A3 - A320F"
GP4_pos(42) = "M1 - M115"
GP4_pos(43) = "M1 - M117"
GP4_pos(44) = "A1 - A114"
GP4_pos(45) = " A1 - A115"
GP4_pos(46) = " F1 - F107"
GP4_pos(47) = " M1 - M114"
GP4_pos(48) = " M1 - M118"
GP4_pos(49) = "A2 - A209"
GP4_pos(50) = " A2 - A210"
GP4_pos(51) = " A2 - A211"
GP4_pos(52) = " A3 - A309"
GP4_pos(53) = " A3 - A310"
GP4_pos(54) = " A3 - A311"
GP4_pos(55) = " M3 - M304"
GP4_pos(56) = " M3 - M305"
GP4_pos(57) = " M4 - M405"
GP4_pos(58) = "F1 - F108"
GP4_pos(59) = " M3 - M308"
GP4_pos(60) = " M4 - M408"
GP4_pos(61) = "M3 - M303"
GP4_pos(62) = " M4 - M403"
GP4_pos(63) = "M3 - M302"
GP4_pos(64) = " M4 - M402"
GP4_pos(65) = "A3 - A313"
GP4_pos(66) = "M2 - M201F"
GP4_pos(67) = " M2 - M201R"
GP4_pos(68) = "M2 - M204F"
GP4_pos(69) = "A2 - A221F"
GP4_pos(70) = " A2 - A222"
GP4_pos(71) = "M2 - M214"
GP4_pos(72) = "A2 - A224F"
GP4_pos(73) = " A2 - A225"
GP4_pos(74) = " A3 - A325"
GP4_pos(75) = "A2 - A223F"
GP4_pos(76) = "M2 - M215F"
GP4_pos(77) = "M2 - M216F"
GP4_pos(78) = "M2 - M213F"
GP4_pos(79) = " M2 - M218F"
GP4_pos(80) = "M2 - M217"
GP4_pos(81) = "A3 - A321F"
GP4_pos(82) = " A3 - A322"
GP4_pos(83) = " A3 - A323F"
GP4_pos(84) = "A3 - A324F"
GP4_pos(85) = "M2 - M208"
GP4_pos(86) = "F1 - F101"
GP4_pos(87) = "M2 - M209"
GP4_pos(88) = " M2 - M211"
GP4_pos(89) = "A2 - A215"
GP4_pos(90) = " A2 - A216"
GP4_pos(91) = " A2 - A217"
GP4_pos(92) = " A3 - A316"
GP4_pos(93) = " A3 - A317"
GP4_pos(94) = " A3 - A318"
GP4_pos(95) = "A2 - A214"
GP4_pos(96) = "M2 - M210"
GP4_pos(97) = "A3 - A302"
GP4_pos(98) = "A3 - A301"
GP4_pos(99) = "A2 - A203"
GP4_pos(100) = " A3 - A303"
GP4_pos(101) = "A2 - A201"
GP4_pos(102) = "M2 - M203F"
GP4_pos(103) = "A1 - A101"
GP4_pos(104) = " A1 - A104"
GP4_pos(105) = "A1 - A103"
GP4_pos(106) = " M2 - M202R"
GP4_pos(107) = " M2 - M205R"
GP4_pos(108) = "A1 - A106"
GP4_pos(109) = "M1 - M101"
GP4_pos(110) = "M1 - M102"
GP4_pos(111) = "A1 - A102"
GP4_pos(112) = "A2 - A204"
GP4_pos(113) = " A2 - A205"
GP4_pos(114) = " A2 - A206"
GP4_pos(115) = " A3 - A304"
GP4_pos(116) = " A3 - A305"
GP4_pos(117) = " A3 - A306"
GP4_pos(118) = "A3 - A314R"
GP4_pos(119) = "A2 - A202"
GP4_pos(120) = " A3 - A314F"
GP4_pos(121) = " A3 - A315F"
GP4_pos(122) = "A3 - A315R"
GP4_pos(123) = "M2 - M207"
GP4_pos(124) = "M2 - M212"
GP4_pos(125) = "A2 - A208"
GP4_pos(126) = " A3 - A308"
GP4_pos(127) = "A2 - A207"
GP4_pos(128) = "A3 - A307"
GP4_pos(129) = "M3 - M301"
GP4_pos(130) = "M4 - M401"
GP4_pos(131) = "A2 - CNTR"
GP4_pos(132) = " F1 - CNTR"
GP4_pos(133) = " M1 - CNTR"
GP4_pos(134) = " M2 - CNTR"
GP4_pos(135) = " M3 - CNTR"
GP4_pos(136) = " M4 - CNTR"
GP4_pos(137) = "F1 - F103"
GP4_pos(138) = "FRS - 503"
GP4_pos(139) = "ZA - SEAT"
GP4_pos(140) = "A3 - A313"
GP4_pos(141) = "A3 - CNTR"
GP4_pos(142) = "F1 - F104"



Dim CBASE_POS(0 To 207) As String

CBASE_POS(0) = "F1 - 01F T"
CBASE_POS(1) = "F1 - 01F M"
CBASE_POS(2) = "F1 - 01F B"
CBASE_POS(3) = "F1 - 03"
CBASE_POS(4) = "F1 - 04"
CBASE_POS(5) = "F1 - 06F"
CBASE_POS(6) = "F1 - 07"
CBASE_POS(7) = "F1 - 08"
CBASE_POS(8) = "F1 - FLTDK"
CBASE_POS(9) = "F1 - ST-1C 1"
CBASE_POS(10) = "F1 - ST-1C 2"
CBASE_POS(11) = "F1 - ST-1C 3"
CBASE_POS(12) = "M1 - 01"
CBASE_POS(13) = "M1 - 02"
CBASE_POS(14) = "M1 - 03"
CBASE_POS(15) = "M1 - 04"
CBASE_POS(16) = "M1 - 05"
CBASE_POS(17) = "M1 - 06"
CBASE_POS(18) = "M1 - 07"
CBASE_POS(19) = "M1 - 08"
CBASE_POS(20) = "M1 - 09"
CBASE_POS(21) = "M1 - 10"
CBASE_POS(22) = "M1 - 11"
CBASE_POS(23) = "M1 - 14"
CBASE_POS(24) = "M1 - 15"
CBASE_POS(25) = "M1 - 16"
CBASE_POS(26) = "M1 - 17"
CBASE_POS(27) = "M1 - 18"
CBASE_POS(28) = "M2 - 01F"
CBASE_POS(29) = "M2 - 01R"
CBASE_POS(30) = "M2 - 02F"
CBASE_POS(31) = "M2 - 02R"
CBASE_POS(32) = "M2 - 03F"
CBASE_POS(33) = "M2 - 03R"
CBASE_POS(34) = "M2 - 04F"
CBASE_POS(35) = "M2 - 04R"
CBASE_POS(36) = "M2 - 05F"
CBASE_POS(37) = "M2 - 05R"
CBASE_POS(38) = "M2 - 06F"
CBASE_POS(39) = "M2 - 06R"
CBASE_POS(40) = "M2 - 07F 1"
CBASE_POS(41) = "M2 - 07R 1"
CBASE_POS(42) = "M2 - 07F 2"
CBASE_POS(43) = "M2 - 07R 2"
CBASE_POS(44) = "M2 - 07F 3"
CBASE_POS(45) = "M2 - 07R 3"
CBASE_POS(46) = "M2 - 07F 4"
CBASE_POS(47) = "M2 - 07R 4"
CBASE_POS(48) = "M2 - 07F 5"
CBASE_POS(49) = "M2 - 07R 5"
CBASE_POS(50) = "M2 - 08"
CBASE_POS(51) = "M2 - 09"
CBASE_POS(52) = "M2 - 10"
CBASE_POS(53) = "M2 - 11"
CBASE_POS(54) = "M2 - 12F 1"
CBASE_POS(55) = "M2 - 12R 1"
CBASE_POS(56) = "M2 - 12F 2"
CBASE_POS(57) = "M2 - 12R 2"
CBASE_POS(58) = "M2 - 12F 3"
CBASE_POS(59) = "M2 - 12R 3"
CBASE_POS(60) = "M2 - 12F 4"
CBASE_POS(61) = "M2 - 12R 4"
CBASE_POS(62) = "M2 - 12F 5"
CBASE_POS(63) = "M2 - 12R 5"
CBASE_POS(64) = "M2 - 13F"
CBASE_POS(65) = "M2 - 14"
CBASE_POS(66) = "M2 - 15F"
CBASE_POS(67) = "M2 - 16F"
CBASE_POS(68) = "M2 - 17"
CBASE_POS(69) = "M2 - 18F"
CBASE_POS(70) = "M3 - 01F 1"
CBASE_POS(71) = "M3 - 01R 1"
CBASE_POS(72) = "M3 - 01F 2"
CBASE_POS(73) = "M3 - 01R 2"
CBASE_POS(74) = "M3 - 01F 3"
CBASE_POS(75) = "M3 - 01R 3"
CBASE_POS(76) = "M3 - 02"
CBASE_POS(77) = "M3 - 03"
CBASE_POS(78) = "M3 - 04"
CBASE_POS(79) = "M3 - 06F"
CBASE_POS(80) = "M3 - 07F"
CBASE_POS(81) = "M3 - 08"
CBASE_POS(82) = "M4 - 01F 1"
CBASE_POS(83) = "M4 - 01R 1"
CBASE_POS(84) = "M4 - 01F 2"
CBASE_POS(85) = "M4 - 01R 2"
CBASE_POS(86) = "M4 - 01F 3"
CBASE_POS(87) = "M4 - 01R 3"
CBASE_POS(88) = "M4 - 02"
CBASE_POS(89) = "M4 - 03"
CBASE_POS(90) = "M4 - 05"
CBASE_POS(91) = "M4 - 06F"
CBASE_POS(92) = "M4 - 07F"
CBASE_POS(93) = "M4 - 08"
CBASE_POS(94) = "A1 - 01"
CBASE_POS(95) = "A1 - 02"
CBASE_POS(96) = "A1 - 03"
CBASE_POS(97) = "A1 - 04"
CBASE_POS(98) = "A1 - 05"
CBASE_POS(99) = "A1 - 06"
CBASE_POS(100) = "A1 - 07"
CBASE_POS(101) = "A1 - 08"
CBASE_POS(102) = "A1 - 09"
CBASE_POS(103) = "A1 - 10"
CBASE_POS(104) = "A1 - 11"
CBASE_POS(105) = "A1 - 12"
CBASE_POS(106) = "A1 - 13"
CBASE_POS(107) = "A1 - 14"
CBASE_POS(108) = "A1 - 15"
CBASE_POS(109) = "A1 - 16"
CBASE_POS(100) = "A2 - 01"
CBASE_POS(111) = "A2 - 02"
CBASE_POS(112) = "A2 - 03"
CBASE_POS(113) = "A2 - 04"
CBASE_POS(114) = "A2 - 05"
CBASE_POS(115) = "A2 - 06"
CBASE_POS(116) = "A2 - 07F 1"
CBASE_POS(117) = "A2 - 07R 1"
CBASE_POS(118) = "A2 - 07F 2"
CBASE_POS(119) = "A2 - 07R 2"
CBASE_POS(120) = "A2 - 07F 3"
CBASE_POS(121) = "A2 - 07R 3"
CBASE_POS(122) = "A2 - 07F 4"
CBASE_POS(123) = "A2 - 07R 4"
CBASE_POS(124) = "A2 - 07F 5"
CBASE_POS(125) = "A2 - 07R 5"
CBASE_POS(126) = "A2 - 08F 1"
CBASE_POS(127) = "A2 - 08R 1"
CBASE_POS(128) = "A2 - 08F 2"
CBASE_POS(129) = "A2 - 08R 2"
CBASE_POS(130) = "A2 - 08F 3"
CBASE_POS(131) = "A2 - 08R 3"
CBASE_POS(132) = "A2 - 08F 4"
CBASE_POS(133) = "A2 - 08R 4"
CBASE_POS(134) = "A2 - 08F 5"
CBASE_POS(135) = "A2 - 08R 5"
CBASE_POS(136) = "A2 - 09"
CBASE_POS(137) = "A2 - 11"
CBASE_POS(138) = "A2 - 14"
CBASE_POS(139) = "A2 - 15"
CBASE_POS(130) = "A2 - 16"
CBASE_POS(141) = "A2 - 17"
CBASE_POS(142) = "A2 - 19F"
CBASE_POS(143) = "A2 - 20F"
CBASE_POS(144) = "A2 - 21F"
CBASE_POS(145) = "A2 - 22"
CBASE_POS(146) = "A2 - 23F"
CBASE_POS(147) = "A2 - 24F"
CBASE_POS(148) = "A2 - 25"
CBASE_POS(149) = "A2 - 26F"
CBASE_POS(150) = "A2 - 28 T"
CBASE_POS(151) = "A2 - 28 M"
CBASE_POS(152) = "A2 - 28 B"
CBASE_POS(153) = "A2 - 29 T"
CBASE_POS(154) = "A2 - 29 M"
CBASE_POS(155) = "A2 - 29 B"
CBASE_POS(156) = "A3 - 01"
CBASE_POS(157) = "A3 - 02"
CBASE_POS(158) = "A3 - 03"
CBASE_POS(159) = "A3 - 04"
CBASE_POS(160) = "A3 - 05"
CBASE_POS(161) = "A3 - 06"
CBASE_POS(162) = "A3 - 07F 1"
CBASE_POS(163) = "A3 - 07F 2"
CBASE_POS(164) = "A3 - 07F 3"
CBASE_POS(165) = "A3 - 07F 4"
CBASE_POS(166) = "A3 - 07F 5"
CBASE_POS(167) = "A3 - 07R 1"
CBASE_POS(168) = "A3 - 07R 2"
CBASE_POS(169) = "A3 - 07R 3"
CBASE_POS(170) = "A3 - 07R 4"
CBASE_POS(171) = "A3 - 07R 5"
CBASE_POS(172) = "A3 - 08F 1"
CBASE_POS(173) = "A3 - 08F 2"
CBASE_POS(174) = "A3 - 08F 3"
CBASE_POS(175) = "A3 - 08F 4"
CBASE_POS(176) = "A3 - 08F 6"
CBASE_POS(177) = "A3 - 08R 1"
CBASE_POS(178) = "A3 - 08R 2"
CBASE_POS(179) = "A3 - 08R 3"
CBASE_POS(180) = "A3 - 08R 4"
CBASE_POS(181) = "A3 - 08R 5"
CBASE_POS(182) = "A3 - 09"
CBASE_POS(183) = "A3 - 11"
CBASE_POS(184) = "A3 - 13"
CBASE_POS(185) = "A3 - 14F"
CBASE_POS(186) = "A3 - 14R"
CBASE_POS(187) = "A3 - 15F"
CBASE_POS(188) = "A3 - 15R"
CBASE_POS(189) = "A3 - 16"
CBASE_POS(190) = "A3 - 17"
CBASE_POS(191) = "A3 - 18"
CBASE_POS(192) = "A3 - 19F"
CBASE_POS(193) = "A3 - 20F"
CBASE_POS(194) = "A3 - 21F"
CBASE_POS(195) = "A3 - 22"
CBASE_POS(196) = "A3 - 23F"
CBASE_POS(197) = "A3 - 24F"
CBASE_POS(198) = "A3 - 25"
CBASE_POS(199) = "A3 - 26F"
CBASE_POS(200) = "A3 - 28 B"
CBASE_POS(201) = "A3 - 28 M"
CBASE_POS(202) = "A3 - 28 T"
CBASE_POS(203) = "A3 - 29 B"
CBASE_POS(204) = "A3 - 29 M"
CBASE_POS(205) = "A3 - 29 T"
CBASE_POS(206) = "A3 - CNTR"
CBASE_POS(207) = "FWD - 503"







End Sub



Sub createXML()

'in ac file add sheet1 and sheet2 copy and paste csv file into sheet1 for code

Dim Sheet1 As Worksheet
Set Sheet1 = Worksheets("Sheet1")

Dim Sheet2 As Worksheet
Set Sheet2 = Worksheets("Sheet2")

    Dim i As Long
    Dim j As Long
    Dim var As String
    
    Dim k As Long
    

    Sheet1.Cells(1, 1).Select
    startrow1 = Sheet1.Cells(1, 1).Row
    Selection.End(xlDown).Select
    endrow1 = Selection.Row
    

    Sheet1.Cells(2, 1).Select
    startrow1 = Sheet1.Cells(2, 1).Row
    Selection.End(xlRight).Select
    endrow1 = Selection.Row
    
    
    For i = startrow1 To endrow2
    For j = startrow2 To endrow2
    
        
    
    
    Next j
    Next i
    
    
    


End Sub
