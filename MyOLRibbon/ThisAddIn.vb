﻿Imports Microsoft.Office.Interop

Public Class ThisAddIn
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New amanda()
    End Function
    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
