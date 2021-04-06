VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private dbsubsec   As clsSubSec
Attribute dbsubsec.VB_VarHelpID = -1
Private Xcor As Integer
Private Ycor As Integer
Private Mode As String
Private System1 As String
Private System2 As String
Private SysCoord1 As String
Private SysCoord2 As String
Private CD1Name As String, CD1Coor As String
Private CD2Name As String, CD2Coor As String

Public Sub makemap2(ByVal mobjPicture As Object)
Dim icolumn As Integer
Dim icount As Integer
Dim y2 As Integer
Dim x2 As Integer
Dim y1 As Integer
Dim x1 As Integer
Dim y3 As Integer
Dim x3 As Integer
Dim irow As Integer
Dim pi As Double

    pi = 4 * Atn(1)
    mobjPicture.ScaleMode = 1
    mobjPicture.Line (0, 0)-Step(mobjPicture.Width, mobjPicture.Height), , BF
    mobjPicture.CurrentX = 300
    mobjPicture.CurrentY = 300

    For icolumn = 1 To 8
        For icount = 30 To 390 Step 60
            y2 = 300 * Cos(icount * (pi / 180))
            x2 = 300 * Sin(icount * (pi / 180))
            If icount = 90 Then
                x1 = mobjPicture.CurrentX
                y1 = mobjPicture.CurrentY
            End If
            If icount = 150 Then
                x3 = mobjPicture.CurrentX
                y3 = mobjPicture.CurrentY
            End If
            mobjPicture.Line Step(0, 0)-Step(x2, y2), 16700000
        Next icount
        
        For irow = 1 To 9
            mobjPicture.CurrentX = x1
            mobjPicture.CurrentY = y1
            For icount = -30 To 360 Step 60
                y2 = 300 * Cos(icount * (pi / 180))
                x2 = 300 * Sin(icount * (pi / 180))
                If icount = 90 Then
                    x1 = mobjPicture.CurrentX
                    y1 = mobjPicture.CurrentY
                End If
                mobjPicture.Line Step(0, 0)-Step(x2, y2), 16700000  '16777215
            Next icount
        Next irow
        
        If icolumn Mod 2 = 1 Then
            mobjPicture.CurrentX = x3
            mobjPicture.CurrentY = y3
        Else
            mobjPicture.CurrentX = x3
            mobjPicture.CurrentY = 300
        End If
        
    Next icolumn
   
End Sub
Public Sub place(ByVal mobjPicture As Object, x As Integer, y As Integer, color As Double, Optional warning As Double)
Dim x2 As Integer
    
    mobjPicture.FillStyle = 0

    Select Case x
        Case Is < 3
            x2 = 1
        Case Is < 5
            x2 = 2
        Case Is < 7
            x2 = 3
        Case Is < 9
            x2 = 4
    End Select
    
    If x Mod 2 = 1 Then
        If warning > 0 Then
            mobjPicture.FillColor = warning
            mobjPicture.Circle (Int(x2) * 900 - 300, (520 * y) - 220), 150, vbBlack
        End If
        mobjPicture.FillColor = color
        mobjPicture.Circle (Int(x2) * 900 - 300, (520 * y) - 220), 75, vbBlack
    Else
        If warning > 0 Then
            mobjPicture.FillColor = warning
            mobjPicture.Circle (Int(x2) * 900 + 150, (520 * y) + 55), 150, vbBlack
        End If
        mobjPicture.FillColor = color
        mobjPicture.Circle (Int(x2) * 900 + 150, (520 * y) + 55), 75, vbBlack
    End If
    
End Sub
Public Sub Placeroute(ByRef mobjPicture As Object, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, color As Double)
Dim iX As Integer
Dim iX2 As Integer
Dim iY As Integer
Dim iY2 As Integer
Dim stx As Integer
Dim stY As Integer
Dim enX As Integer
Dim enY As Integer
  
    mobjPicture.FillStyle = 0
    mobjPicture.FillColor = color
    
    Select Case x1
        Case Is < -1
            iX = -1
        Case Is < 1
            iX = 0
        Case Is < 3
            iX = 1
        Case Is < 5
            iX = 2
        Case Is < 7
            iX = 3
        Case Is < 9
            iX = 4
        Case Is < 11
            iX = 5
        Case Is < 13
            iX = 6
    End Select
    
    Select Case x2
        Case Is < -1
            iX2 = -1
        Case Is < 1
            iX2 = 0
        Case Is < 3
            iX2 = 1
        Case Is < 5
            iX2 = 2
        Case Is < 7
            iX2 = 3
        Case Is < 9
            iX2 = 4
        Case Is < 11
            iX2 = 5
        Case Is < 13
            iX2 = 6
    End Select
    
    If Abs(x1 Mod 2) = 1 Then
       stx = iX * 900 - 300
       stY = y1 * 520 - 220
       
    Else
       stx = iX * 900 + 150
       stY = y1 * 520 + 55
    End If
    
    If Abs(x2 Mod 2) = 1 Then
       enX = iX2 * 900 - 300
       enY = y2 * 520 - 220
    Else
       enX = iX2 * 900 + 150
       enY = y2 * 520 + 55
    End If
    mobjPicture.DrawWidth = 2
   ' mobjPicture.DrawStyle = 2
    mobjPicture.Line (enX, enY)-(stx, stY), color
    mobjPicture.DrawWidth = 1
   ' mobjPicture.DrawStyle = 0

End Sub
Public Sub DetectLocation(ByRef mobjPicture As Object, x As Single, y As Single)
Dim i As Integer
Dim XI As Integer
Dim YI As Integer
Dim y2 As Double
Dim x2 As Double
Dim y1 As Double
Dim x1 As Double
Dim Ydiff As Double
Dim Distance As Double
Dim mCoor As Object
    If mobjPicture.combo1.Text = "" Then Exit Sub
    XI = ((x - 150) / 450)
    If XI Mod 2 = 1 Then y = y + 210
    YI = (y / 520)

    For i = 0 To mobjPicture.Controls.Count - 1
    
    'Debug.Print mobjPicture.Controls(i).Name & "  " & i & TypeName(mobjPicture.Controls(i))
        
        If TypeOf mobjPicture.Controls(i) Is TextBox Then
            mobjPicture.Controls(i).Text = ""
        End If
        If TypeName(mobjPicture.Controls(i)) = "RichTextBox" Then
            mobjPicture.Controls(i).Text = ""
        End If
        
    Next i

    If YI > 0 And YI < 11 And XI > 0 And XI < 9 Then
        mobjPicture.TxtCoord = Right("00" & (XI + Xcor - 1), 2) & Right("00" & (YI + Ycor - 1), 2)
    End If
    
    Set mCoor = dbsubsec.GetRecordset("SELECT SubSec.*, System.* From subSec, System WHERE CoX = " & XI + Xcor - 1 & " and Coy = " & YI + Ycor - 1 & ";")
    
    If mCoor.EOF = False Then
        mobjPicture.txtName = mCoor.Fields("SysName")
        mobjPicture.txtCr = mCoor.Fields("CR") & ""
        mobjPicture.txtTL = mCoor.Fields("TL") & ""
        mobjPicture.txtwtn = mCoor.Fields("WTN") & ""
        
        mobjPicture.txtstrprt = mCoor.Fields("StrPrt") & ""
        mobjPicture.txtdiam = mCoor.Fields("diameter") & ""
        mobjPicture.txtgrav = mCoor.Fields("gravity") & ""
        mobjPicture.txtatmos = mCoor.Fields("Atmos") & ""
        mobjPicture.txtMoon = mCoor.Fields("Moons") & ""
        mobjPicture.txtorbit = mCoor.Fields("orbit") & ""
        mobjPicture.txtpb = mCoor.Fields("pb") & ""
        mobjPicture.txtgov = mCoor.Fields("GOV") & ""
        mobjPicture.txtHydro = mCoor.Fields("Hydro") & ""
        'mobjPicture.txtalleguanc = mCoor.Fields("alleguanc") & ""
        
        mobjPicture.txtstrprt = mCoor.Fields("strprt") & ""
        mobjPicture.txtclim = mCoor.Fields("clim") & ""
        mobjPicture.txtpop = mCoor.Fields("pop") & ""
        mobjPicture.txtgg = mCoor.Fields("GG") & ""
        
        mobjPicture.txtstrprt = mCoor.Fields("strprt") & ""
        
        
        Set mCoor = dbsubsec.GetRecordset("SELECT * From SystemNotes WHERE SystemName='" & mobjPicture.txtName & "';")
        
        Do Until mCoor.EOF
            mobjPicture.txtnotes.Text = mobjPicture.txtnotes.Text & mCoor.Fields("line")
            mCoor.MoveNext
        Loop
    End If
    Set mCoor = Nothing
    
    If Mode = "CD1" Then
        If mobjPicture.txtName.Text > "" Then
            CD1Name = mobjPicture.txtName.Text
        Else
            CD1Name = "Coord: " & mobjPicture.TxtCoord
        End If
        CD1Coor = mobjPicture.TxtCoord
        Mode = "CD2"
        MsgBox "Select Ending Location"
    Else
        If Mode = "CD2" Then
            If mobjPicture.txtName.Text > "" Then
                CD2Name = mobjPicture.txtName.Text
            Else
                CD2Name = "Coord: " & mobjPicture.TxtCoord
            End If
            CD2Coor = mobjPicture.TxtCoord
            
            x1 = CInt(Left(CD1Coor, 2))
            x2 = CInt(Left(CD2Coor, 2))
            
            y1 = CInt(Right(CD1Coor, 2))
            y2 = CInt(Right(CD2Coor, 2))
            
            Distance = CheckDistance(x1, y1, x2, y2)

      
            MsgBox "The Distance from " & CD1Name & " to " & CD2Name & " is " & Distance
        
            Mode = "N"
        End If
    End If
    
End Sub

Public Sub LoadMap(ByRef mobjPicture As Object, Subsect As String)
Dim mCoor As Object
Dim mobjRSRoute As Object
Dim warning As Double
Dim color As Double
    If Subsect = "" Then Exit Sub
    makemap2 mobjPicture
  
    Set mobjRSRoute = dbsubsec.GetRecordset("SELECT Route.*, SubSec.* FROM Route, SubSec WHERE Route.SSName=SubSec.SSName and SubSec.SSName = '" & Subsect & "';")
    
    Do Until mobjRSRoute.EOF
        Select Case mobjRSRoute.Fields("Type")
            Case "X"
                color = vbGreen
            Case "DX"
                color = vbYellow
            Case "ZPR"
                color = vbRed
        End Select
        Placeroute mobjPicture, (mobjRSRoute.Fields("Xco1") - mobjRSRoute.Fields("Xcor") + 1), (mobjRSRoute.Fields("Yco1") - mobjRSRoute.Fields("Ycor") + 1), (mobjRSRoute.Fields("Xco2") - mobjRSRoute.Fields("Xcor") + 1), (mobjRSRoute.Fields("Yco2") - mobjRSRoute.Fields("Ycor") + 1), color
        mobjRSRoute.MoveNext
    Loop
    
    Set mCoor = dbsubsec.GetRecordset("SELECT SubSec.*, System.* From subSec, System WHERE SubSec.SSName=System.SubSec and SSName = '" & Subsect & "';")
    
    If Not mCoor.EOF Then
        Xcor = mCoor.Fields("Xcor")
        Ycor = mCoor.Fields("Ycor")
    End If
    
    Do Until mCoor.EOF
        warning = 0
        If mCoor.Fields("Zone") = "Red" Then warning = vbRed
        If mCoor.Fields("Zone") = "Amber" Then warning = vbYellow
        place mobjPicture, mCoor.Fields("CoX") - mCoor.Fields("Xcor") + 1, mCoor.Fields("CoY") - mCoor.Fields("Ycor") + 1, vbWhite, warning
        mCoor.MoveNext
    Loop
    If Mode = "" Then Mode = "N"
End Sub
Public Sub LoadSubsector(ByRef mobjPicture As Object)
Dim col As Object
    Set col = New Collection
    Set dbsubsec = New clsSubSec
    Set col = dbsubsec.GetRecordset("select SSName from subsec")
    col.MoveFirst
    
    Do Until col.EOF = True
        mobjPicture.AddItem col.Fields(0)
        col.MoveNext
    Loop

    Set col = Nothing
End Sub
Public Sub CalcDistance(mstrSubSector As String)
    If Mode <> "N" Then Exit Sub
    If Mode = "N" Then
        Mode = "CD1"
        MsgBox "Select Starting Location"
    End If
End Sub

Private Sub Class_Initialize()
'    Mode = "N"
End Sub

Private Function CheckDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
Dim xs As Double
Dim ys As Double
Dim partDistance As Integer
    
    If y1 = y2 And x1 = x2 Then
        CheckDistance = 0
        Exit Function
    End If

If y1 = y2 Then
    If x1 < x2 Then
        xs = x1 + 1
        If xs = x2 Then
            CheckDistance = 1
            Exit Function
        Else
            partDistance = CheckDistance(xs, y1, x2, y2)
            CheckDistance = partDistance + 1
            Exit Function
        End If
    End If
    If x1 > x2 Then
        xs = x2 + 1
        If xs = x1 Then
            CheckDistance = 1
            Exit Function
        Else
            partDistance = CheckDistance(x1, y1, xs, y2)
            CheckDistance = partDistance + 1
            Exit Function
        End If
    End If
End If

If x1 = x2 Then
    If y1 < y2 Then
        ys = y1 + 1
        If ys = y2 Then
            CheckDistance = 1
            Exit Function
        Else
            partDistance = CheckDistance(x1, ys, x2, y2)
            CheckDistance = partDistance + 1
            Exit Function
        End If
    End If
    If y1 > y2 Then
        ys = y2 + 1
        If ys = y1 Then
            CheckDistance = 1
            Exit Function
        Else
            partDistance = CheckDistance(x1, y1, x2, ys)
            CheckDistance = partDistance + 1
            Exit Function
        End If
    End If
End If

If x1 < x2 And y1 < y2 Then
        xs = x1 + 1
        ys = y1 + 1
    If xs = x2 And ys = y2 Then
        CheckDistance = 1 + x1 Mod 2
        Exit Function
    Else
        partDistance = CheckDistance(xs, ys, x2, y2)
        CheckDistance = 1 + x1 Mod 2
        Exit Function
    End If
End If

If x1 > x2 And y1 > y2 Then
        xs = x2 + 1
        ys = y2 + 1
    If xs = x1 And ys = y1 Then
        CheckDistance = 1 + x2 Mod 2
        Exit Function
    Else
        partDistance = CheckDistance(x1, y1, xs, ys)
        CheckDistance = 1 + x2 Mod 2
        Exit Function
    End If
End If

If x1 < x2 And y1 > y2 Then
        xs = x1 + 1
        ys = y2 + 1
    If xs = x2 And ys = y1 Then
        CheckDistance = 1 + x1 Mod 2
        Exit Function
    Else
        partDistance = CheckDistance(xs, y1, x2, ys)
        CheckDistance = 1 + x1 Mod 2
        Exit Function
    End If
End If

If x1 > x2 And y1 < y2 Then
        xs = x2 + 1
        ys = y1 + 1
    If xs = x1 And ys = y2 Then
        CheckDistance = 1 + x2 Mod 2
        Exit Function
    Else
        partDistance = CheckDistance(x1, ys, xs, y2)
        CheckDistance = 1 + x2 Mod 2
        Exit Function
    End If
End If

End Function
