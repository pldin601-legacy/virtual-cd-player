Attribute VB_Name = "Default_Subs"
Declare Function GetTickCount Lib "kernel32.dll" () As Long

Global IDXPosIn(32000) As Long
Global IDXPosOut(32000) As Long
Function Trim32(inString As String)

InSrting = Trim(inString)

For X = Len(inString) To 1 Step -1
 If Asc(Mid(inString, X, 1)) > 32 Then inString = Mid(inString, 1, X): Exit For
Next

For X = 1 To Len(inString) Step 1
 If Asc(Mid(inString, X, 1)) < 32 Then Trim32 = Mid(inString, 1, X - 1): Exit Function
Next

Trim32 = inString

End Function

Function Scaler(sA As Long, sB As Long, sStep As Long, sSteps As Long) As Long

Dim sC As Long

sC = sB - sA

Scaler = sC + (sA / sSteps * sStep)

End Function



Sub Задержка(Миллисекунд As Long)

Dim ВремЗнач As Long

ВремЗнач = GetTickCount

Do: DoEvents: Loop While Not GetTickCount - ВремЗнач > Миллисекунд

End Sub

Function MMod(mLong)
  If mLong < 0 Then MMod = -mLong Else MMod = mLong
End Function

Function MaxVal(inVal1, inVal2)
If inVal1 > inVal2 Then MaxVal = inVal1 Else MaxVal = inVal2
End Function

Function MaxValue(ByRef inVal1 As Long, ByRef inVal2 As Long) As Long

If inVal1 > inVal2 Then MaxValue = inVal1 Else MaxValue = inVal2

End Function

Function CountChars(inChar As String, inString As String) As Integer
k = 0
For X = 1 To Len(inString)
 If Mid(inString, X, 1) = inChar Then k = k + 1
Next
CountChars = k
End Function

Function EnPass(inText) As String
Dim inTmp, X

inTmp = String(Len(inText), 32)

For X = 1 To Len(inText)
 Mid(inTmp, X, 1) = Chr(255 - Asc(Mid(inText, X, 1)))
Next

EnPass = inTmp

End Function


Function GetLongFromData(InDay As Integer, InMonth As Integer, InYear As Integer) As Integer

If InYear Mod 4 = 0 Then N = 1 Else N = 0
If InMonth = 1 Then temp = InDay
If InMonth = 2 Then temp = InDay + 31
If InMonth = 3 Then temp = InDay + 31 + 28 + N
If InMonth = 4 Then temp = InDay + 31 + 28 + 31 + N
If InMonth = 5 Then temp = InDay + 31 + 28 + 31 + 30 + N
If InMonth = 6 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + N
If InMonth = 7 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + N
If InMonth = 8 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + N
If InMonth = 9 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + N
If InMonth = 10 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + N
If InMonth = 11 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + N
If InMonth = 12 Then temp = InDay + 31 + 28 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + N
GetLongFromData = temp
End Function

Sub GetDateFromLong(InLong As Integer, InDay As Integer, InMonth As Integer, InYear As Integer)

Dim temp As Integer, A As Integer, B As Integer, N As Integer

If InYear Mod 4 = 0 Then N = 1 Else N = 0
temp = InLong
If temp > 0 Then A = temp: B = 1
If temp > 31 Then A = temp - 31: B = 2
If temp > 59 + N Then A = temp - 59 - N: B = 3
If temp > 90 + N Then A = temp - 90 - N: B = 4
If temp > 120 + N Then A = temp - 120 - N: B = 5
If temp > 151 + N Then A = temp - 151 - N: B = 6
If temp > 181 + N Then A = temp - 181 - N: B = 7
If temp > 212 + N Then A = temp - 212 - N: B = 8
If temp > 243 + N Then A = temp - 243 - N: B = 9
If temp > 273 + N Then A = temp - 273 - N: B = 10
If temp > 304 + N Then A = temp - 304 - N: B = 11
If temp > 334 + N Then A = temp - 334 - N: B = 12

InDay = A
InMonth = B

End Sub


Function GetFileSize(FName As String) As Long
 On Error Resume Next
 I = FreeFile
 Open FName For Input As #I
 GetFileSize = LOF(I)
 Close I
End Function


Function GetTimeFromMinutes(vMinutes As Long)
If vMinutes < 60 * 60 Then GetTimeFromMinutes = Format$(Fix(vMinutes / 60), "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
If vMinutes >= 60 * 60 Then GetTimeFromMinutes = Format$(Fix(vMinutes / 3600), "00") & ":" & Format$(Fix(vMinutes / 60) Mod 60, "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
End Function

Function GetMinutesFromTime(vTime As String)
Dim MinS, Hors
Hors = Val(Mid$(vTime, 1, 2))
MinS = Val(Mid(vTime, 4, 2))
GetMinutesFromTime = (Hors * 60) + MinS
End Function

Public Function GetVersion() As String
GetVersion = Format$(App.Major, "0") + "." + Format$(App.Minor, "0") + "." + Format$(App.Revision, "000")
End Function

Public Function Get2Version() As String
Get2Version = Format$(App.Major, "0") + "." + Format$(App.Minor, "0")
End Function

Function PathHead$(FileName As String)
Dim Names As Integer
For Names = Len(FileName) To 1 Step -1
 If Mid$(FileName, Names, 1) = "\" Then
  PathHead$ = Mid$(FileName, 1, (Names) - 1)
  If PathHead$ = "$APPDIR$" Then PathHead$ = App.Path
  Exit For
 End If
Next

End Function

Function FileExists(Path$) As Boolean
    Dim X As Integer

    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X
    Err.Clear

End Function

Public Function FileHead$(FileName As String)
Dim Names As Integer
For Names = Len(FileName) To 1 Step -1
If Mid$(FileName, Names, 1) = "\" Then FileHead$ = Right$(FileName, Len(FileName) - (Names)): Exit Function
Next
End Function

Public Function LowPath(InPath As String) As String
If Right$(InPath, 1) = "\" Then LowPath = InPath
If Right$(InPath, 1) <> "\" Then LowPath = InPath + "\"
End Function

Public Function GetIniRecord(Record As String, INIFile As String, Optional rDefault = "") As String
Dim CfgLine As String, G As Integer
On Error Resume Next
G = FreeFile
Open INIFile For Input As #G
Do
Line Input #G, CfgLine
If UCase$(Mid$(CfgLine, 1, Len(Record))) = UCase(Record) Then
   GetIniRecord = Mid$(CfgLine, Len(Record) + 1)
   Close G: Exit Function
End If
Loop While Not EOF(G)
GetIniRecord = Format(rDefault)
Close G
End Function

Public Function Загрузить_Настройки2(Опция As String, Имя_Файла As String, Optional По_Умолчанию = "") As String
Dim CfgLine As String, G As Integer
On Error Resume Next
G = FreeFile
Open Имя_Файла For Input As #G
Do
Line Input #G, CfgLine
If UCase$(Right$(CfgLine, Len(Опция))) = UCase(Опция) Then
   Загрузить_Настройки2 = Mid$(CfgLine, 1, Len(CfgLine) - Len(Опция) - 1)
   Close G: Exit Function
End If
Loop While Not EOF(G)
Загрузить_Настройки2 = Format(По_Умолчанию)
Close G
End Function

Public Function GetNetCard(Record As String)
Dim CfgLine As String, G As Integer
G = InStr(1, Record, " - ")
If G > 0 Then
 CfgLine = Trim(Mid(Record, 1, G))
Else
 CfgLine = Record
End If

GetNetCard = CfgLine
End Function

Function ReadCommand(ByRef GetCommand As String, ByRef GetValue As Boolean)
 If GetValue = True Then ReadCommand = Right$(GetCommand, Len(GetCommand) - 12)
 If GetValue = False Then ReadCommand = Mid$(GetCommand, 1, 11)
End Function

Function FilterName(Text As String) As String

Dim LS, Bs, Variants, Bizer
On Error Resume Next

For LS = 1 To Len(Text)
Bs = Mid$(Text, LS, 1)

 For Variants = 0 To 47
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 91 To 96
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 58 To 63
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 123 To 191
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 
Mid$(Text, LS, 1) = Bs

Next

If Text = "" Then Text = "Unnamed"
FilterName = Text

End Function

Function CBol(Value) As Boolean
If Not Val(Format(Value)) = 0 Then CBol = True: Exit Function
If Format(Value) = "True" Then CBol = True: Exit Function
CBol = False
End Function

Function xStr(Value As Boolean) As String
If Value = True Then xStr = "True" Else xStr = "False"
End Function

Function Bol2Int(inVal As Boolean) As Integer
  Bol2Int = 0
  If inVal = True Then Bol2Int = 1
  If inVal = False Then Bol2Int = 0
End Function


Sub ExchangeFiles(SI As Integer, DI As Integer, Sources As ListBox)
On Error Resume Next
Dim A, B, ASel As Boolean, BSel As Boolean
A = Sources.List(DI)
B = Sources.List(SI)
ASel = Sources.Selected(DI)
BSel = Sources.Selected(SI)
Sources.List(DI) = B
Sources.List(SI) = A
Sources.Selected(DI) = BSel
Sources.Selected(SI) = ASel

End Sub

' Sub Sleep(TM)
' tm1 = Timer
' Do: DoEvents: Loop While Not Timer >= tm1 + TM
' End Sub

Sub ExtractData(inFile As String, outFile As String, fstByte As Long, lenByte As Long)
On Error Resume Next


Open inFile For Binary As #11
Open outFile For Binary As #12

Const BUFLEN = 32666
LessData = lenByte Mod BUFLEN
OkedData = Fix(lenByte / BUFLEN)
Dim BufferString As String * BUFLEN

For N = fstByte To lenByte + 1 Step BUFLEN
 Get #11, N, BufferString
 Put #12, , BufferString
Next N

Close #11, #12

End Sub
