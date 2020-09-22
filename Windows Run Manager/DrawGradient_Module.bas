Attribute VB_Name = "DrawGradient_Module"
'Function DrawGradient "v.3.0", Made by Derek Tremblay -----------------------

Public Function DrawGradient(TheObj As Object, ColorRed%, ColorGreen%, _
                             ColorBlue%, ColorStop%, ColorBandSize%, _
                             StartLine%, StopLine%, ModLine%, _
                             Optional AutoReDrawObj As Boolean = True, _
                             Optional MoreR% = 1, Optional MoreG% = 1, _
                             Optional MoreB% = 1)
  On Error Resume Next

    Dim sngBlueCur As Single, sngRedCur As Single, sngGreenCur As Single
    Dim sngBlueStep As Single, sngRedStep As Single, sngGreenStep As Single
    Dim intFormHeight As Integer, intFormWidth As Integer, intY As Integer
  '
  'Initialize Variable MoreX -------------------------------------------------
  '
      If MoreR% <= 0 Then MoreR% = 1
      If MoreG% <= 0 Then MoreG% = 1
      If MoreB% <= 0 Then MoreB% = 1
    
      If MoreR% >= 20 Then MoreR% = 20
      If MoreG% >= 20 Then MoreG% = 20
      If MoreB% >= 20 Then MoreB% = 20
  '
  'Make Object "AutoRedraw" --------------------------------------------------
  '
    Select Case AutoReDrawObj
      Case True
        TheObj.AutoRedraw = True
      Case False
        TheObj.AutoRedraw = False
    End Select
  '
  'Get system values for height and width ------------------------------------
  '
    intFormHeight = TheObj.ScaleHeight
    intFormWidth = TheObj.ScaleWidth
  '
  'Calculate step size and Color start value ---------------------------------
  '
    sngRedStep = ColorBandSize% * (ColorStop% - ColorRed%) / intFormHeight
    sngRedCur = ColorRed%
    '
    sngGreenStep = ColorBandSize% * (ColorStop% - ColorGreen%) / intFormHeight
    sngGreenCur = ColorGreen%
    '
    sngBlueStep = ColorBandSize% * (ColorStop% - ColorBlue%) / intFormHeight
    sngBlueCur = ColorBlue%
  '
  'Paint Color screen --------------------------------------------------------
  '
    For intY = StartLine% To StopLine% Step ColorBandSize%
       TheObj.Line (-1, intY - 1)-(intFormWidth, intY + ColorBandSize% _
                   \ ModLine%), RGB(sngRedCur \ MoreR%, sngGreenCur \ MoreG% _
                   , sngBlueCur \ MoreB%), BF

        sngBlueCur = sngBlueCur + sngBlueStep
        sngRedCur = sngRedCur + sngRedStep
        sngGreenCur = sngGreenCur + sngGreenStep
    Next intY
  '
  '---------------------------------------------------------------------------
  '
End Function

