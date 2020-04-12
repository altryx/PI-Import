Module mdlMain

    Sub Main()
        On Error GoTo errhandler
        '-op ave -tag tag1, tag2 -ts "2017-01-01" -te "2017-02-01" -td 1h
        Call process_cmdline_arguments()
        Call process_options()

        Select Case PI_Cmd
            Case "cv"
                For Each Tag In PI_Tags
                    Dim curr_value As PISDK.PIValue = get_current_value(Tag)
                    Console.WriteLine(IIf(output_tagname, Tag & output_sep, "") &
                                      IIf(output_timestamp, Format(curr_value.TimeStamp.LocalDate, "G") & output_sep, "") &
                                      Str(curr_value.Value))
                Next
            Case "summary"
                For Each Tag In PI_Tags
                    Dim summary_values As PISDK.PIValues = get_sum_value(PI_Summary_Operation, Tag, PI_StartTime, PI_EndTime, PI_Duration)
                    Dim summary_val As Integer
                    For summary_val = 1 To summary_values.Count 
                        Console.WriteLine(IIf(output_tagname, Tag & output_sep, "") &
                                          IIf(output_timestamp, Format(summary_values(summary_val).ValueAttributes("EarliestTime").Value.localdate, "G") & output_sep, "") &
                                          IIf(summary_values(summary_val).ValueAttributes(1).Name Like "Err*", " ", CStr(summary_values(summary_val).Value)))

                    Next
                Next
            Case "hv"
                For Each Tag In PI_Tags
                    Dim historical_values As PISDK.PIValues = get_hist_val(PI_Tag, PI_HistTimes)
                    Dim hist_val As Integer
                    For hist_val = 1 To historical_values.Count
                        Console.WriteLine(IIf(output_tagname, Tag & output_sep, "") &
                                          IIf(output_timestamp, Format(historical_values(hist_val).TimeStamp.LocalDate, "G") & output_sep, "") &
                                          Str(historical_values(hist_val).Value))

                    Next
                Next
        End Select

        Exit Sub
errhandler:
        Console.WriteLine("ERROR " & Str(Err.Number) & " : " & Err.Description)
        End

    End Sub

   
End Module
