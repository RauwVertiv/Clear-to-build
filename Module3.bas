Attribute VB_Name = "Module3"
' ============================================
' MÓDULO 3: Activación del evento global
' ============================================

Public XAppEvents As clsAppEvents

Public Sub ActivarEventoGlobal()
    Set XAppEvents = New clsAppEvents
    Set XAppEvents.App = Application
End Sub

Public Sub DesactivarEventoGlobal()
    ' Restaurar formato antes de desactivar
    If Not prevCell Is Nothing Then
        RestaurarFormatoOriginal prevCell
        Set prevCell = Nothing
    End If
    Set XAppEvents = Nothing
End Sub
