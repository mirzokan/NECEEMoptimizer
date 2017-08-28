Option Explicit

Sub optimize()
    Dim max_itterations As Integer, i As Integer
    Dim solved As Boolean, fast_sep_pref As Boolean
    Dim lim_code As String, talker As String
  
    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual
    ' Application.DisplayAlerts = False 

    sh_data.Calculate
    sh_report.Calculate
    sh_naaped.Calculate
    sh_calcval.Calculate

    max_itterations = sh_settings.Range("max_itterations").Value

    ' Check Maximization preference
    If sh_data.Range("pref_separation").Value = sh_settings.Range("A2").Value Then
        fast_sep_pref = True
    Else
        fast_sep_pref = False
    End If

    If check_validated() = True and fast_sep_pref = True Then
        talker = "Great! Your current results are validated and are trustworthy!"
    ElseIf sh_calcval.Range("check_loq").Value = "FAIL" Then
        talker = "Sorry, but your instrument does not have a sufficient limit of quantitation to properly measure such a low Kd value. You will need to find a way to improve the sensitivity of the method. Try increasing the diameter of the capillary."
    Else
        For i = 0 To max_itterations
            sh_data.Calculate
            sh_report.Calculate
            sh_naaped.Calculate
            sh_calcval.Calculate
            sh_calcval.Range("iter_rec").Value = sh_calcval.Range("iter_rec").Value +1
         'Check for success        
         solved = check_solved()
            If fast_sep_pref = True Then 'Short Analysis
                If solved = True Then
                    talker = "Hooray! While your current results are not validated, we've found a set of conditions which will make things better!"
                    Exit For
                End If
            Else 'Maximum separation 
                If solved = True and check_max() = True Then
                    talker = "Hooray! While your current results are not validated, we've found a set of conditions which will make things better, with the best possible separation too!"
                    Exit For
                End If
            End If    

            ' When things don't go so well
            ' Check what the E limits are
            lim_code = check_E_interval()
            If InStr(lim_code, "F") <> 0 Then
                talker = "There are problems with the software settings. Please download the latest version of the file."
                Exit For
            End If

            Select Case lim_code
                Case "VV","DC","DT","DD"
                    Exit For
                Case "VC","VT","VD","DV"
                    Call caplen_up
                    fast_sep_pref = False
                Case "XV","XC","XT","XD","IV","IC","IT","ID"
                    Call seplen_down
            End Select          
        Next i
    End If    
    sh_report.Range("talker:F31").ClearContents
    If check_solved() = False and check_validated() = False and sh_calcval.Range("check_loq").Value = "PASS" Then
        talker = talk_speaker(lim_code)
    End If

    If check_validated = True and sh_calcval.Range("iter_rec").Value > 0 Then
        talker = "Hooray! Your current results are validated, though we've found a set of conditions which can improve the separation!"
    End If
        sh_report.Range("talker:F31").Value = talker

    ' Application.ScreenUpdating = True
    ' Application.Calculation = xlCalculationAutomatic
    ' Application.DisplayAlerts = True
End Sub

Function check_validated() As Boolean
    sh_data.Calculate
    sh_report.Calculate
    sh_naaped.Calculate
    sh_calcval.Calculate
    If sh_calcval.Range("validated").Value = "PASS" Then
        check_validated = True
    Else
        check_validated = False
    End If
End Function

Function check_solved() As Boolean
    sh_data.Calculate
    sh_report.Calculate
    sh_naaped.Calculate
    sh_calcval.Calculate
    If sh_calcval.Range("solved").Value = "PASS" Then
        check_solved = True
    Else
        check_solved = False
    End If
End Function

Function check_max() As Boolean
    If sh_calcval.Range("maxed").Value = "MAXED" Then
        check_max = True
    Else
        check_max = False
    End If
End Function

Function check_E_interval() as String
    Dim minE_limit As String, maxE_limit As String, combination As String, maxE_code As String, minE_code As String

    minE_limit = sh_calcval.Range("ind_minE").Value
    Select Case minE_limit
        Case "Voltage Limited"
            minE_code = "V"
        Case "Complex dissociation limited"
            minE_code = "X"
        Case "Ion depletion time limited"
            minE_code = "I"
        Case "tau limited"
            minE_code = "D"
        Case "FAIL"
            minE_code = "F"
    End Select

    maxE_limit = sh_calcval.Range("ind_maxE").Value
    Select Case maxE_limit
        Case "Voltage Limited"
            maxE_code = "V"
        Case "Current limited"
            maxE_code = "C"
        Case "Temperature limited"
            maxE_code = "T"
        Case "tau limited"
            maxE_code = "D"
        Case "FAIL"
            maxE_code = "F"
    End Select

    check_E_interval = minE_code & maxE_code
End Function

Sub seplen_down()
    If sh_calcval.Range("pred_prop_compx").Value = "Cap" Then
        sh_calcval.Range("sepgoal_ovrr").Value = sh_calcval.Range("sepgoal_iter").Value
    Else
        sh_calcval.Range("prop_ovrr").Value = sh_calcval.Range("prop_iter").Value
    End If
End Sub

Sub caplen_up()
    sh_calcval.Range("sepgoal_ovrr").Value = sh_calcval.Range("sep_max").Value
End Sub

Function talk_speaker(code As String) As String
    Select Case code
        Case "VV"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. It seems that the instrument information has been entered incorrectly. Please, verify that maximum and minumum voltages have been entered correctly."
        Case "VC"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high current. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "VT"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high temperature. Consider allowing a higher temperature, or please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "VD"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Even at the smallest possible electric field strength (smallest velocities) the complex dissociation is too slow to produce a detectable decay region. Try injecting a wider plug of the sample, but ideally you need to perform the experiment in a much longer capillary than the current instrument allows."
        Case "XV"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Even at the highest possible electric field strength (largest velocities) the complex moves too slowly to reach the detector before fully dissociating. In other words, the complex dissociates faster than it can be efficiently separated from the ligand by CE. NECEEM might not be the best method for this interaction."
        Case "XC"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high current. Due to the limited electric field strength, he complex moves too slowly to reach the detector before fully dissociating. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "XT"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high temperature. Due to the limited electric field strength, he complex moves too slowly to reach the detector before fully dissociating. Consider using a higher temperature, or try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "XD"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Dissociation is too slow, but the complex dissociates before it reaches the detector. This is an unusual case, and is most likely caused by improper use of this program. Please download the latest version of this file and try again. If this message repeats, please send us an e-mail, we would love to take a look at this strange case!"
        Case "IV"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Even at the highest possible electric field strength (largest velocities) the analytes move too slowly to reach the detector before buffer depletion sets on. Try increasing the volume or the buffering capacity of the run buffer to allow for more analysis time."
        Case "IC"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high current. Due to the limited electric field strength, the analytes move too slowly to reach the detector before buffer depletion sets on. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "IT"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high temperature. Due to the limited electric field strength, the analytes move too slowly to reach the detector before buffer depletion sets on. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "ID"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Complex dissociates too slowly, and takes too long to elute. Try increasing the volume or the buffering capacity of the run buffer to allow for more analysis time."
        Case "DV"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Even at the highest possible electric field strength (fastest velocities) the separation is too slow to match the high rate of dissociation. This results in a decay ragion that is too pronounced and is difficult to deconvolute. Try injecting a narrower plug of the sample, but ideally you need to find a way to increase the difference in mobilities of the complex and the ligand."
        Case "DC"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high current. Due to the limited electric field strength, the separation between the species is too slow to match the high rate of dissociation. This results in a decay ragion that is too pronounced and is difficult to deconvolute. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "DT"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. Cannot reach optimization due to high temperature. Due to the limited electric field strength, the separation between the species is too slow to match the high rate of dissociation. This results in a decay ragion that is too pronounced and is difficult to deconvolute. Please try to reduce the conductivity of your system, either by using a capillary of a smaller diameter, or reduce the ionic strength of the running buffer."
        Case "DD"
            talk_speaker ="Sorry, we could not optimize the experimental conditions. There are problems with the software settings. Please download the latest version of the file."
    End Select
End Function
