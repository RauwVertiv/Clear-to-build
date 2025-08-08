VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ResaltarCeldaActiva Target
End Sub

