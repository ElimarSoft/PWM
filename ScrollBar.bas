Attribute VB_Name = "ScrollBar"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
 
Sub ScrollBar1_Change()

    Dim ws1 As Worksheet
    
    Set ws1 = Sheets("Graphics")
    Dim value As Double: value = ws1.Shapes("Scroll Bar 1").ControlFormat.value
    
    ws1.Range("AY15").value = 100 - value

    Dim index As Double
    index = 1 - value / 100
    Update (index)

End Sub

Public Sub AutoTest()

    Dim n As Integer
    
    For n = 0 To 100 Step 10
        Update (n / 100)
        Dim ws1 As Worksheet
        Set ws1 = ActiveSheet
        DoEvents
        Application.Wait (Now + TimeValue("0:00:10"))
        'Sleep (10000)
    Next n

End Sub
