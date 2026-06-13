VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Declare the exact types of controls we are grouping together
Public WithEvents HoverFrame As MSForms.Frame
Attribute HoverFrame.VB_VarHelpID = -1
Public WithEvents HoverLabel As MSForms.Label
Attribute HoverLabel.VB_VarHelpID = -1
Public WithEvents HoverIcon As MSForms.Image
Attribute HoverIcon.VB_VarHelpID = -1

Private ParentForm As Object

' 2. Initialize now accepts all three control types
Public Sub Initialize(ByVal ctrlFrame As MSForms.Frame, ByVal ctrlLabel As MSForms.Label, ByVal ctrlIcon As MSForms.Image, ByVal frm As Object)
    Set HoverFrame = ctrlFrame
    Set HoverLabel = ctrlLabel
    Set HoverIcon = ctrlIcon
    Set ParentForm = frm
End Sub

' 3. Listen for the mouse moving over ANY of the three items
Private Sub HoverFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TriggerHoverEffect
End Sub

Private Sub HoverLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TriggerHoverEffect
End Sub

Private Sub HoverIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TriggerHoverEffect
End Sub

' 4. The shared hover effect
Private Sub TriggerHoverEffect()
    ' NEW: If the panel just got clicked, ignore the mouse!
    If ParentForm.DisableHover = True Then Exit Sub

    ParentForm.BtnColorReset
    
    ' Highlight the frame and the label using your custom Hoverblue constant
    HoverFrame.BackColor = Hoverblue
    HoverLabel.BackColor = Hoverblue
    
    ParentForm.PannelSizeBig
End Sub
