' VBAProject1
'IN VBA unlock VBProject with password and delete all modules in it.


Sub VBPRO1()
Dim vbProj As Object
Dim co As Variant
Set vbProj = ThisWorkbook.VBProject
'unlock
  If vbProj.Protection <> 1 Then
'delete modules  
                        For Each co In vbProj.vbcomponents
                        On Error Resume Next
                        vbProj.vbcomponents.Remove vbProj.vbcomponents(co.Name)
               
  Else
            Set Application.VBE.ActiveVBProject = vbProj
            
'Replace VBPassword with your current VBProject password            
            
            SendKeys "VBPassword" & "~~"
            
            Application.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute
            
            Application.DisplayAlerts = False
                        'Ignore errors
            On Error Resume Next
  End If

'delete modules
                        For Each co In vbProj.vbcomponents
                        On Error Resume Next
                        vbProj.vbcomponents.Remove vbProj.vbcomponents(co.Name)


End sub
