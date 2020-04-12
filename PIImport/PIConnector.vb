Module PIConnector
    Public PIObject As New PISDK.PISDK
    Public PIConn As PISDK.Server

    Sub define_connection(PI_Server)
        On Error GoTo errhandler
        If PI_Server Like "default" Then
            PIConn = PIObject.Servers.DefaultServer
        Else
            PIConn = PIObject.Servers(PI_Server)
        End If

        Exit Sub
errhandler:
        Console.WriteLine("ERROR " & Str(Err.Number) & " : " & Err.Description)
        End
    End Sub

    Function get_current_value(PI_Tag) As PISDK.PIValue
        get_current_value = PIConn.PIPoints(PI_Tag).Data.Snapshot
    End Function

    Function get_hist_val(PI_Tag, PI_StartTime) As PISDK.PIValues
        Dim PIData As PISDK.IPIData2 = PIConn.PIPoints(PI_Tag).Data
        get_hist_val = PIData.TimedValues(PI_StartTime)
    End Function

    Function get_sum_value(PI_Operation, PI_Tag, PI_StartTime, PI_EndTime, PI_Duration) As PISDK.PIValues
        Dim PIData As PISDK.IPIData2 = PIConn.PIPoints(PI_Tag).Data
        Dim return_summary_value As PISDKCommon.NamedValues
        return_summary_value = PIData.Summaries2(PI_StartTime, PI_EndTime, PI_Duration, PI_Operation, PISDK.CalculationBasisConstants.cbTimeWeighted)
        get_sum_value = return_summary_value(ret_PI_Op_Descriptor(PI_Operation)).Value
    End Function
    Function ret_PI_Op_Descriptor(PI_Operation)
        Select Case PI_Operation
            Case PISDK.ArchiveSummariesTypeConstants.asMaximum
                ret_PI_Op_Descriptor = "Maximum"
            Case PISDK.ArchiveSummariesTypeConstants.asMinimum
                ret_PI_Op_Descriptor = "Minimum"
            Case PISDK.ArchiveSummariesTypeConstants.asCount
                ret_PI_Op_Descriptor = "Count"
            Case PISDK.ArchiveSummariesTypeConstants.asAverage
                ret_PI_Op_Descriptor = "Average"
            Case PISDK.ArchiveSummariesTypeConstants.asTotal
                ret_PI_Op_Descriptor = "Total"
            Case PISDK.ArchiveSummariesTypeConstants.asRange
                ret_PI_Op_Descriptor = "Range"
            Case PISDK.ArchiveSummariesTypeConstants.asStdDev
                ret_PI_Op_Descriptor = "StdDev"
            Case Else
                ret_PI_Op_Descriptor = ""
        End Select

    End Function
End Module
