VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim cControl As CommandBarButton


Private Sub Workbook_AddinInstall()


On Error Resume Next 'Just in case

    'Delete any existing menu item that may have been left.

    Application.CommandBars("Worksheet Menu Bar").Controls("Sales Forecast").Delete

    'Add the new menu item and Set a CommandBarButton Variable to it

    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add

    'Work with the Variable

        With cControl

            .Caption = "Sales Forecast"

            .FaceId = 1845
            
            .Style = msoButtonCaption

            .OnAction = "Reformat"

            'Macro stored in a Standard Module

        End With

        

    On Error GoTo 0


End Sub
Private Sub Workbook_AddinUninstall()

    

On Error Resume Next 'In case it has already gone.

    Application.CommandBars("Worksheet Menu Bar").Controls("Super Code").Delete

    On Error GoTo 0

End Sub
