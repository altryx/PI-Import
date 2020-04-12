Module mdlUtil
    Public PI_Tag As String = String.Empty
    Public PI_Tags As String()
    Public PI_StartTime As String = String.Empty
    Public PI_HistTimes As String()
    Public PI_EndTime As String = String.Empty
    Public PI_Duration As String = String.Empty
    Public PI_Cmd As String = "cv"
    Public PI_Summary_Operation As PISDK.ArchiveSummariesTypeConstants
    Public PI_Operation_Descriptor As String
    Public PI_Server As String = "default"
    Public output_csv As Boolean = False
    Public output_csv_file As String = String.Empty
    Public output_stdout As Boolean = False
    Public output_db As Boolean = False
    Public output_db_name As String = String.Empty
    Public output_tagname As Boolean = False
    Public output_timestamp As Boolean = True
    Public output_sep As String = ","

    Sub process_cmdline_arguments()
        Dim cmdargs() As String = Environment.GetCommandLineArgs()
        Dim cmd_arg As Integer

        If UBound(cmdargs) > 1 Then
            For cmd_arg = 1 To cmdargs.Count - 1 Step 2
                If cmdargs(cmd_arg) = "-op" Then
                    PI_Cmd = cmdargs(cmd_arg + 1)
                    Select Case LCase(PI_Cmd)
                        Case "ave", "average"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asAverage
                        Case "min", "minimum"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asMinimum
                        Case "max", "maximum"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asMaximum
                        Case "count"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asCount
                        Case "sum", "total"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asTotal
                        Case "stdev"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asStdDev
                        Case "range", "rng"
                            PI_Cmd = "summary"
                            PI_Summary_Operation = PISDK.ArchiveSummariesTypeConstants.asRange
                        Case "curr", "current", "cv"
                            PI_Cmd = "cv"
                        Case "hist", "histval", "hv"
                            PI_Cmd = "hv"
                        Case Else
                            Console.WriteLine("Missing or invalid summary operation.")
                            End
                    End Select
                End If

                If cmdargs(cmd_arg) = "-tag" Then
                    PI_Tag = cmdargs(cmd_arg + 1)
                    If PI_Tag = "" Or PI_Tag Like "-*" Then
                        Console.WriteLine("Missing or invalid PI Tag supplied.")
                        End
                    Else
                        PI_Tags = Split(PI_Tag, ",")
                        If UBound(PI_Tags) > 0 Then output_tagname = True
                    End If
                End If

                If cmdargs(cmd_arg) = "-ts" Then
            PI_StartTime = cmdargs(cmd_arg + 1)
            If PI_StartTime = "" Or PI_StartTime Like "-*" Then
                Console.WriteLine("Missing or invalid PI Start Time.")
                End
            End If
        End If

        If cmdargs(cmd_arg) = "-times" Then
            Dim hist_times As String
            hist_times = cmdargs(cmd_arg + 1)
            If hist_times = "" Or hist_times Like "-*" Then
                Console.WriteLine("Missing or invalid list of times for historical data retreival.")
                End
            ElseIf Not (PI_Cmd = "hv") Then
                Console.WriteLine("Parameter -times should only be used with -op hist|hv|histval switch")
                End
            Else
                PI_HistTimes = Split(hist_times, ",")
            End If
        End If

        If cmdargs(cmd_arg) = "-te" Then
            PI_EndTime = cmdargs(cmd_arg + 1)
            If PI_EndTime = "" Or PI_EndTime Like "-*" Then
                Console.WriteLine("Missing or invalid PI End Time.")
                End
            End If
        End If

        If cmdargs(cmd_arg) = "-td" Then
            PI_Duration = cmdargs(cmd_arg + 1)
            If PI_Duration = "" Or PI_Duration Like "-*" Then
                Console.WriteLine("Missing or invalid PI Duration.")
                End
            End If
        End If

        If cmdargs(cmd_arg) = "-server" Then
            PI_Server = cmdargs(cmd_arg + 1)
            If PI_Server = "" Or PI_Tag Like "-*" Then
                Console.WriteLine("Missing or invalid PI Server supplied.")
                End
            End If
        End If

        If cmdargs(cmd_arg) = "-out_csv" Then output_csv = True : output_csv_file = cmdargs(cmd_arg + 1)
        If cmdargs(cmd_arg) = "-out_stdout" Then output_stdout = True
        If cmdargs(cmd_arg) = "-out_db" Then output_db = True : output_db_name = cmdargs(cmd_arg + 1)
        If cmdargs(cmd_arg) = "-incl_tagname" Then output_tagname = True
        If cmdargs(cmd_arg) = "-excl_timestamp" Then output_timestamp = False

        If cmdargs(cmd_arg) = "-sep" Then
            output_sep = cmdargs(cmd_arg + 1)
            If output_sep = "" Or output_sep Like "-*" Then
                Console.WriteLine("Missing or invalid output separator.")
                End
            End If
        End If
            Next
        Else
            Call print_app_header()
            Console.WriteLine(
                              "Usage:" & vbCrLf &
                              "-tag <tag_name1>,[tag_name2],... PI point or a comma-separated list of points" & vbCrLf &
                              "-op [current, value_at_time, average, sum, min, max, ] - PI Data operation to undertake" & vbCrLf &
                              "-ts start_time - PI compatible start-time definition" & vbCrLf &
                              "-te end_time - PI compatible end-time definition" & vbCrLf &
                              "-td duration - PI compatible definition of duration (only summary calcs)" & vbCrLf &
                              "-server server_name - PI Data server to use [if omitted use default]" & vbCrLf &
                              "")
            End
        End If



       


       

        'Console.WriteLine(PI_Operation & " " & PI_Op)
        'Command line argument structure
        '
        'INPUTS
        '-tag <tag_name> Tag Name
        '-op [current, value_at_time, average, sum, min, max, ] - PI Data operation to undertake
        '-ts start_time - PI compatible start-time definition
        '-te end_time - PI compatible end-time definition
        '-dur duration - PI compatible definition of duration (only summary calcs)
        '-server server_name - PI Data server to use [if omitted use default]
        '
        'OUTPUTS
        '-out_csv <csv_file_path> - Output to a CSV file provided in the path
        '-out_stdout - Output to standard output
        '-out_db <Connecction profile>- Output to a database defined by connection profile (see help)
        '
    End Sub
    Sub process_options()
        PIConnector.define_connection(PI_Server)
    End Sub
    Sub print_app_header()
        Console.WriteLine("PI Import " & My.Application.Info.Version.ToString & " [" & My.Application.Info.Version.Build.ToString & "] " & vbCrLf &
                           "Damir Lampa 2018" & vbCrLf)
    End Sub
End Module
