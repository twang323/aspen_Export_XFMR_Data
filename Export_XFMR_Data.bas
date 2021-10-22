' ASPEN Script
' Export Transformer Data1.BAS
'
	REM Begin Dialog Dialog_1 49,60,260,123, "ASPEN Update Line Impedance"
	  REM OptionGroup .GROUP_1
		REM OptionButton 81,17,28,8, "TLC"
		REM OptionButton 125,17,28,8, "TCIM"
	  REM OptionGroup .GROUP_2
		REM OptionButton 135,54,28,8, "69 kV"
		REM OptionButton 135,70,36,8, "138 kV"
	  REM Text 10,16,60,8,"New Line Data:"
	  REM Text 58,49,23,8,"Rate A:"
	  REM Text 58,63,23,8,"Rate B:"
	  REM Text 58,78,23,8,"Rate C:"
	  REM Text 10,49,42,8,"MVA Rating"
	  REM TextBox 10,29,242,12,.EditBox_data
	  REM OKButton 57,99,48,12
	  REM CancelButton 138,99,48,12
	  REM TextBox 89,47,34,11,.RateA
	  REM TextBox 89,61,34,11,.RateB
	  REM TextBox 89,76,34,11,.RateC
	REM End Dialog

Sub main()
  'Dim dlg As Dialog_1
  Dim Rating(4) As Double
  Dim Amp(4) As Double
  vnDate$ = Date()
  'dlg.GROUP_1 = 0
  'dlg.EditBox_data = ""
  'If 0 = Dialog(dlg) Then Exit Sub
  

  'newData$ = dlg.EditBox_data
  'EditBox_1 = "U:\Programming\VB\CableZ VB Program\ExportTransformer.txt"
  EditBox_1 = "ExportTransformer.txt"
  Open EditBox_1 For Output As #1

  
  ' ===================== New Line Data From TLC (Follow the Example Format)==============================================
  'TLC = 1
  '  newData$ = "Text80:	13.323	Text84:	0.01517	Text85:	0.05427	Text81:	0.01406	Text82:	0.04887	Text83:	0.13841	Text103: 0.00934"
  
  'newData$ = "Text80:	6.0210	Text84:	0.0720	Text85:	0.0995	Text81:	0.0016	Text82:	0.1083	Text83:	0.3113	Text103: 0.0008"
  
  ' ===================== End of Data Entry ================================================

  '-------- End of New Data Update------------

  'If GetData( HND_SYS, SY_dBaseMVA, BaseMVA ) = 0 Then BaseMVA = 100
  ' Figure out kV base from picked object
  If 0 <> GetEquipment( TC_PICKED, PickedHnd ) Then
    ' Probe to see what's being picked
    Select Case EquipmentType( PickedHnd )
      Case TC_XFMR3
	    If 0 = GetData( PickedHnd, X3_dTap1, T_kV1#  ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dTap2, T_kV2#  ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dTap3, T_kV3#  ) Then GoTo HasError
        If 0 = GetData( PickedHnd, X3_dRps, T_Rps# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, X3_dXps, T_Xps# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, X3_dRpt, T_Rpt#  ) Then GoTo HasError
        If 0 = GetData( PickedHnd, X3_dXpt, T_Xpt#  ) Then GoTo HasError
        If 0 = GetData( PickedHnd, X3_dRst, T_Rst#) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dXst, T_Xst# ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dR0ps, T_R0ps# ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dX0ps, T_X0ps#) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dR0pt, T_R0pt#  ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dX0pt, T_X0pt# ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dR0st, T_R0st# ) Then GoTo HasError
		If 0 = GetData( PickedHnd, X3_dX0st, T_X0st#  ) Then GoTo HasError
		'If 0 = GetData( nBusHnd1,BUS_sName,nBus1$) Then GoTo HasError
		'If 0 = GetData( nBusHnd2,BUS_sName,nBus2$) Then GoTo HasError
		'If 0 = GetData( PickedHnd, LN_vdRating, Rating) Then GoTo HasError
		'aLine1$ = "'"& GetObjMemo (PickedHnd)&"' /"
		'Equipment$= "Line impedance: "
      Case Else
	  
        Print "Please select a transformer"
        exit Sub
    End Select
 End If   
	  '-------- Start of New Rating Update--------

  'BaseZ = BaseKV * BaseKV / BaseMVA
  'dX  = dX * BaseZ
  'dX0 = dX0 * BaseZ
  'dR  = dR * BaseZ
  'dR0 = dR0* BaseZ
  'aLine$ = Equipment$ & "Z1=" & Str(dR) & " +j " & Str(dX) & "  Z0= " & Str(dR0) & " +j " & Str(dX0)

  Print #1, Format(T_kV1,"###0.0") & "/" & Format(T_kV2,"###0.0")& "/" & Format(T_kV3,"###0.0") & " kV transformer"
  Print #1, "Positive Sequence on a 100 MVA Base "
  Print #1, "Zps= (" & Format(T_Rps,"###0.00000") & " +j " & Format(T_Xps,"###0.00000")& _
            "), Zpt=(" & Format(T_Rpt,"###0.00000") & " +j " & Format(T_Xpt,"###0.00000")& _
            "), Zst=(" & Format(T_Rst,"###0.00000") & " +j " & Format(T_Xst,"###0.00000")& ")"
  Print #1, ""
  Print #1, "Zero Sequence on a 100 MVA Base" 
  Print #1, "Zps= (" & Format(T_R0ps,"###0.00000") & " +j " & Format(T_X0ps,"###0.00000")& _
            "), Zpt=(" & Format(T_R0pt,"###0.00000") & " +j " & Format(T_X0pt,"###0.00000")& _
            "), Zst=(" & Format(T_R0st,"###0.00000") & " +j " & Format(T_X0st,"###0.00000")& ")"    
  Print "Data Exported under U:\Programming\VB\CableZ VB Program\ExportTransformer.txt!"
Exit Sub
  ' Error handling
  HasError:
  Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()

