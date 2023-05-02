Attribute VB_Name = "MovePin"

Function SetPinPositionWithoutMovingShape(ByVal Shape As Visio.Shape, ByVal NewPinPosition As String) As Boolean
    ' Declare variables
    Dim OldPinX As Double
    Dim OldPinY As Double
    Dim DeltaX As Double
    Dim DeltaY As Double
    Dim OldLocPinX As Double
    Dim OldLocPinY As Double
    Dim OldLocPinXFormulaU As String
    Dim OldLocPinYFormulaU As String
    Dim Width As Double
    Dim Height As Double
    Dim NewLocPinX As Double
    Dim NewLocPinY As Double

    On Error GoTo ErrorHandler

    ' Get shape's width and height
    Width = Shape.Cells("Width").Result("in")
    Height = Shape.Cells("Height").Result("in")
    
    ' Store the current absolute pin position
    OldPinX = Shape.Cells("PinX").Result("in")
    OldPinY = Shape.Cells("PinY").Result("in")

    ' Store the current local pin position
    OldLocPinX = Shape.Cells("LocPinX").Result("in")
    OldLocPinY = Shape.Cells("LocPinY").Result("in")
    OldLocPinXFormulaU = Shape.Cells("LocPinX").FormulaU
    OldLocPinYFormulaU = Shape.Cells("LocPinY").FormulaU

    ' Determine new pin position based on the input.
    '
    Select Case NewPinPosition
        Case "Top-Left"
            NewLocPinX = 0
            NewLocPinY = Height
            Shape.Cells("LocPinX").FormulaU = "Width*0"
            Shape.Cells("LocPinY").FormulaU = "Height*1"
        Case "Top-Center"
            NewLocPinX = Width / 2
            NewLocPinY = Height
            Shape.Cells("LocPinX").FormulaU = "Width*0.5"
            Shape.Cells("LocPinY").FormulaU = "Height*1"
        Case "Top-Right"
            NewLocPinX = Width
            NewLocPinY = Height
            Shape.Cells("LocPinX").FormulaU = "Width*1"
            Shape.Cells("LocPinY").FormulaU = "Height*1"
        Case "Center-Left"
            NewLocPinX = 0
            NewLocPinY = Height / 2
            Shape.Cells("LocPinX").FormulaU = "Width*0"
            Shape.Cells("LocPinY").FormulaU = "Height*0.5"
        Case "Center-Center"
            NewLocPinX = Width / 2
            NewLocPinY = Height / 2
            Shape.Cells("LocPinX").FormulaU = "Width*0.5"
            Shape.Cells("LocPinY").FormulaU = "Height*0.5"
        Case "Center-Right"
            NewLocPinX = Width
            NewLocPinY = Height / 2
            Shape.Cells("LocPinX").FormulaU = "Width*1"
            Shape.Cells("LocPinY").FormulaU = "Height*0.5"
        Case "Bottom-Left"
            NewLocPinX = 0
            NewLocPinY = 0
            Shape.Cells("LocPinX").FormulaU = "Width*0"
            Shape.Cells("LocPinY").FormulaU = "Height*0"
        Case "Bottom-Center"
            NewLocPinX = Width / 2
            NewLocPinY = 0
            Shape.Cells("LocPinX").FormulaU = "Width*0.5"
            Shape.Cells("LocPinY").FormulaU = "Height*0"
        Case "Bottom-Right"
            NewLocPinX = Width
            NewLocPinY = 0
            Shape.Cells("LocPinX").FormulaU = "Width*1"
            Shape.Cells("LocPinY").FormulaU = "Height*0"
        Case Else
            MsgBox "Invalid pin position provided."
            SetPinPositionWithoutMovingShape = False
            Exit Function
    End Select


    ' Calculate the change in pin position
    DeltaX = NewLocPinX - OldLocPinX
    DeltaY = NewLocPinY - OldLocPinY

    ' Set the new absolute pin position.  (This will move the shape back where it was)
    Shape.Cells("PinX").Result("in") = OldPinX + DeltaX
    Shape.Cells("PinY").Result("in") = OldPinY + DeltaY

    ' Return success
    SetPinPositionWithoutMovingShape = True
    Exit Function

ErrorHandler:
    ' Return failure in case of error
    SetPinPositionWithoutMovingShape = False
End Function


Sub PinBottomLeft()

    Dim VisioApp As Visio.Application
    Dim VisioPage As Visio.Page
    Dim SelectedShape As Visio.Shape
    Dim PinPositions() As String

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    ' Get the Visio application and active page
    Set VisioApp = Application
    Set VisioPage = VisioApp.ActivePage

    ' Check if a shape is selected
    If VisioApp.ActiveWindow.Selection.Count > 0 Then
        Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    Else
        MsgBox "Please select a shape and try again."
    End If
    
    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size & Position 2-D")

    Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    If SetPinPositionWithoutMovingShape(SelectedShape, "Bottom-Left") Then
        Application.EndUndoScope UndoScopeID1, True
    Else
        MsgBox "Error: Could not change pin position."
    End If
    
    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
End Sub

Sub PinBottomRight()

    Dim VisioApp As Visio.Application
    Dim VisioPage As Visio.Page
    Dim SelectedShape As Visio.Shape
    Dim PinPositions() As String

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    ' Get the Visio application and active page
    Set VisioApp = Application
    Set VisioPage = VisioApp.ActivePage

    ' Check if a shape is selected
    If VisioApp.ActiveWindow.Selection.Count > 0 Then
        Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    Else
        MsgBox "Please select a shape and try again."
    End If
    
    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size & Position 2-D")

    Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    If SetPinPositionWithoutMovingShape(SelectedShape, "Bottom-Right") Then
        Application.EndUndoScope UndoScopeID1, True
    Else
        MsgBox "Error: Could not change pin position."
    End If
    
    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
End Sub


Sub PinTopLeft()

    Dim VisioApp As Visio.Application
    Dim VisioPage As Visio.Page
    Dim SelectedShape As Visio.Shape
    Dim PinPositions() As String

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    ' Get the Visio application and active page
    Set VisioApp = Application
    Set VisioPage = VisioApp.ActivePage

    ' Check if a shape is selected
    If VisioApp.ActiveWindow.Selection.Count > 0 Then
        Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    Else
        MsgBox "Please select a shape and try again."
    End If
    
    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size & Position 2-D")

    Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    If SetPinPositionWithoutMovingShape(SelectedShape, "Top-Left") Then
        Application.EndUndoScope UndoScopeID1, True
    Else
        MsgBox "Error: Could not change pin position."
    End If
    
    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
End Sub

Sub PinTopRight()

    Dim VisioApp As Visio.Application
    Dim VisioPage As Visio.Page
    Dim SelectedShape As Visio.Shape
    Dim PinPositions() As String

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    ' Get the Visio application and active page
    Set VisioApp = Application
    Set VisioPage = VisioApp.ActivePage

    ' Check if a shape is selected
    If VisioApp.ActiveWindow.Selection.Count > 0 Then
        Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    Else
        MsgBox "Please select a shape and try again."
    End If
    
    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size & Position 2-D")

    Set SelectedShape = VisioApp.ActiveWindow.Selection.Item(1)
    If SetPinPositionWithoutMovingShape(SelectedShape, "Top-Right") Then
        Application.EndUndoScope UndoScopeID1, True
    Else
        MsgBox "Error: Could not change pin position."
    End If
    
    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
End Sub
