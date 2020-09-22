Attribute VB_Name = "modMain"
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function timeGetTime Lib "winmm.dll" () As Long
Public BenchStart As Long
Public eByte(399) As Byte, EPos As Long

Private Type PatchSaveStructure
Knobs(19) As Integer
Checks(1) As Byte
ComboText(1) As String * 30
End Type: Dim PSS As PatchSaveStructure

Public Function RndRange(ByVal Min As Single, ByVal Max As Single) As Single
'This Function Generates a Random number between 2 numbers.
RndRange = (Rnd * (Max - Min + 1)) + Min
End Function

Sub SavePatch(Filename As String)
PSS.Knobs(0) = Form1.Knob1.KnobValue
PSS.Knobs(1) = Form1.Knob2.KnobValue
PSS.Knobs(2) = Form1.Knob3.KnobValue
PSS.Knobs(3) = Form1.Knob4.KnobValue
PSS.Knobs(4) = Form1.Knob5.KnobValue
PSS.Knobs(5) = Form1.Knob6.KnobValue
PSS.Knobs(6) = Form1.Knob7.KnobValue
PSS.Knobs(7) = Form1.Knob8.KnobValue
PSS.Knobs(8) = Form1.Knob9.KnobValue
PSS.Knobs(9) = Form1.Knob10.KnobValue
PSS.Knobs(10) = Form1.Knob11.KnobValue
PSS.Knobs(11) = Form1.Knob12.KnobValue
PSS.Knobs(12) = Form1.Knob13.KnobValue
PSS.Knobs(13) = Form1.Knob14.KnobValue
PSS.Knobs(14) = Form1.Knob15.KnobValue
PSS.Knobs(15) = Form1.Knob16.KnobValue
PSS.Knobs(16) = Form1.Knob17.KnobValue
PSS.Knobs(17) = Form1.Knob18.KnobValue
PSS.Knobs(18) = Form1.Knob19.KnobValue
PSS.Knobs(19) = Form1.Knob20.KnobValue
PSS.Checks(0) = Form1.EC.Value
PSS.Checks(1) = Form1.WCE.Value
PSS.ComboText(0) = Form1.FX1.Text
PSS.ComboText(1) = Form1.FX2.Text

Open Filename For Binary As #1
Put #1, , PSS
Put #1, , eByte()
Close #1
End Sub

Sub LoadPatch(Filename As String)
Open Filename For Binary As #1
Get #1, , PSS
Get #1, , eByte()
Close #1

Form1.Knob1.SetVal PSS.Knobs(0)
Form1.Knob2.SetVal PSS.Knobs(1)
Form1.Knob3.SetVal PSS.Knobs(2)
Form1.Knob4.SetVal PSS.Knobs(3)
Form1.Knob5.SetVal PSS.Knobs(4)
Form1.Knob6.SetVal PSS.Knobs(5)
Form1.Knob7.SetVal PSS.Knobs(6)
Form1.Knob8.SetVal PSS.Knobs(7)
Form1.Knob9.SetVal PSS.Knobs(8)
Form1.Knob10.SetVal PSS.Knobs(9)
Form1.Knob11.SetVal PSS.Knobs(10)
Form1.Knob12.SetVal PSS.Knobs(11)
Form1.Knob13.SetVal PSS.Knobs(12)
Form1.Knob14.SetVal PSS.Knobs(13)
Form1.Knob15.SetVal PSS.Knobs(14)
Form1.Knob16.SetVal PSS.Knobs(15)
Form1.Knob17.SetVal PSS.Knobs(16)
Form1.Knob18.SetVal PSS.Knobs(17)
Form1.Knob19.SetVal PSS.Knobs(18)
Form1.Knob20.SetVal PSS.Knobs(19)
Form1.EC.Value = PSS.Checks(0)
Form1.WCE.Value = PSS.Checks(1)
Form1.FX1.Text = Trim(PSS.ComboText(0))
Form1.FX2.Text = Trim(PSS.ComboText(1))
End Sub
