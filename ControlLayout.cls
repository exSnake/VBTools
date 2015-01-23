VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AnchorEdges
    LeftAnchor = 2 ^ 0
    TopAnchor = 2 ^ 1
    RightAnchor = 2 ^ 2
    BottomAnchor = 2 ^ 3
    AnchorAll = LeftAnchor + TopAnchor + RightAnchor + BottomAnchor
End Enum

Private anchors As AnchorEdges
Private margins As SideMargins
Private size As ControlSize

Private formSize As ControlSize
Private boundControl As Control

Public Sub Bind(form As UserForm, formControl As Control, anchor As AnchorEdges)
    
    anchors = anchor
    
    formSize.height = form.InsideHeight
    formSize.width = form.InsideWidth
    
    size.height = formControl.height
    size.width = formControl.width
    
    If (anchor And BottomAnchor) = BottomAnchor Then
        margins.BottomMargin = formSize.height - formControl.Top - formControl.height
    End If
    
    If (anchor And LeftAnchor) = LeftAnchor Then
        margins.LeftMargin = formControl.Left
    End If
    
    If (anchor And RightAnchor) = RightAnchor Then
        margins.RightMargin = formSize.width - formControl.Left - formControl.width
    End If
    
    If (anchor And TopAnchor) = TopAnchor Then
        margins.TopMargin = formControl.Top
    End If
    
    Set boundControl = formControl

End Sub

Public Sub Resize(ByVal form As UserForm)
    
    formSize.height = form.InsideHeight
    formSize.width = form.InsideWidth
    
    layout boundControl
    
End Sub

Public Sub layout(ByVal formControl As Control)
    
    If (anchors And TopAnchor) = TopAnchor Then
        
        If (anchors And BottomAnchor) = BottomAnchor Then
            formControl.height = formSize.height - formControl.Top - margins.BottomMargin
        End If
    
    ElseIf (anchors And BottomAnchor) = BottomAnchor Then
        formControl.Top = formSize.height - formControl.height - margins.BottomMargin
    End If
    
    
    If (anchors And LeftAnchor) = LeftAnchor Then
        
        If (anchors And RightAnchor) = RightAnchor Then
            formControl.width = formSize.width - formControl.Left - margins.RightMargin
        End If
    
    ElseIf (anchors And RightAnchor) = RightAnchor Then
        formControl.Left = formSize.width - formControl.width - margins.RightMargin
    End If
    
End Sub
