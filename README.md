' VBAProject1
'IN VBA unlock VBProject with password and delete all modules in it.


Sub VBPRO1()
Dim vbProj As Object
Set vbProj = ThisWorkbook.VBProject
'unlock
  If vbProj.Protection <> 1 Then
'delete modules  
                        For Each element In vbProj.vbcomponents
                        On Error Resume Next
                        vbProj.vbcomponents.Remove element
                        next
               
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
                        For Each element In vbProj.vbcomponents
                        On Error Resume Next
                        vbProj.vbcomponents.Remove element
                        next


End sub
