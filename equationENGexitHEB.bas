Attribute VB_Name = "NewMacros"

Sub InsertEquationENG()
Attribute InsertEquationENG.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.InsertEquationENG"
'
' InsertEquationENG Macro
'
'

Dim objRange As Range
Dim objEq As OMath

Set objRange = Selection.Range
objRange.Text = ""
Set objRange = Selection.OMaths.Add(objRange)
Set objEq = objRange.OMaths(1)
objEq.BuildUp

Application.Keyboard (1033)

End Sub

Sub exitEquationHEB()
Attribute exitEquationHEB.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.exitEquationHEB"
'
' exitEquationHEB Macro
'
'

Selection.EndKey Unit:=wdLine
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.Keyboard (1037)
    Selection.TypeText Text:=" "
    
End Sub
