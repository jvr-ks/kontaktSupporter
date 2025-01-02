; kontaktSXVariables.ahk
; Part of kontaktSupporter.ahk

;---------------------------- initGlobalVariables ----------------------------
initGlobalVariables(){
  global
  
  aquot := "`""
  ablank := " " ;"
  aqb := aquot ablank
  abq := ablank aquot
  aqbq := aquot ablank aquot

  pToken := 0

  fontName := "Segoe UI"
  fontsize := 10
  fontNamePreview := "Consolas"
  fontsizePreview := 11

  tableFilePath := "ksTabellen"
  tableFilePathInUse := tableFilePath
  mTbCols := 9
  ;shorten8 := 15
  selectedTableFile := "ksTabelle.txt"
  autoRepairTableFile := 0

  bodyFilePath := "ksInhalte"
  bodyFilePathInUse := bodyFilePath

  protoFilePath := "ksVorlagen" 
  protoFilePathInUse := protoFilePath
  
  attachmentsPath := "ksAnhang"
  attachmentsPathAbsolute := ""
  attachment := "" ; sub-path of attachmentsPathAbsolute

  textModulsFilePath := "ksTextBausteine"
  textModulsFilePathInUse := textModulsFilePath

  externalAppsFile := "ksExternalApps.txt"
  externalApps := []

  selectedTableFile := "ksTabelle.txt"
  logfileName := "ksGesendet.txt"
  
  substitutionCharLF := Chr(0x00AC) ; "¬"
  separatorChar := Chr(0x00A6) ; "¦"
  
  voiceIsEnabled := 1

  adrTag := "mailto://"
  ccTag := ""
  bccTag := ""
  subjectTag := ""

  mTb := [] ; 2-dimensional
  mTbLine := []

  eMailAppId := 0

  adrText := ""
  ccText := ""
  bccText := ""
  subjectText := ""
  bodyText := ""

  openLink :=  ""
  eMailAppFound := 0

  sendDelay := 200
  currentLineNumber := 0

  ;------------------------------- gui variables -------------------------------
  disableScintillaDefault := 0
  disableScintilla := 0

  minPosLeft := -16
  minPosTop := -16

  minHeightMargin := 100
  minWidthMargin := 100

  mainIsVisible := 0
  LV1 := 0
  menuTxtVisible := 1

  previewIsVisible := 0
  displayPreview := 0
  previewTextMarginHeight := 150

  imagePreviewIsVisible := 0

  quickIsHelpVisible := 0
  rowEditIsVisible := 0

  dpiScaleValueDefault := 96
  dpiScaleValue := dpiScaleValueDefault

  clientTopX := 2
  clientTopY := 2

  LV1MarginTop := clientTopY
  LV1MarginBottom := 25

  LV1MarginLeft := clientTopX
  LV1MarginRight := 2

  if ((0 + A_ScreenDPI == 0) || (A_ScreenDPI == 96))
    dpiCorrect := 1
  else
    dpiCorrect := A_ScreenDPI / dpiScaleValue
    
  guiMainPosXDefault := coordsAppToScreen(10) 
  guiMainPosYDefault := coordsAppToScreen(10)

  guiMainPosX := guiMainPosXDefault
  guiMainPosY := guiMainPosYDefault

  guiMainClientWidthDefault := 600
  guiMainClientHeightDefault := 200

  guiMainClientWidth := guiMainClientWidthDefault
  guiMainClientHeight := guiMainClientHeightDefault


  guiPreviewPosXDefault := coordsAppToScreen(30) 
  guiPreviewPosYDefault := coordsAppToScreen(30)

  guiPreviewPosX := guiPreviewPosXDefault
  guiPreviewPosY := guiPreviewPosYDefault

  guiPreviewClientWidthDefault := 600
  guiPreviewClientHeightDefault := 200

  guiPreviewClientWidth := guiPreviewClientWidthDefault
  guiPreviewClientHeight := guiPreviewClientHeightDefault


  guiImagePreviewPosXDefault := coordsAppToScreen(30) 
  guiImagePreviewPosYDefault := coordsAppToScreen(30)

  guiImagePreviewPosX := guiImagePreviewPosXDefault
  guiImagePreviewPosY := guiImagePreviewPosYDefault

  guiImagePreviewClientWidthDefault := 800
  guiImagePreviewClientHeightDefault := 600

  guiImagePreviewClientWidth := guiImagePreviewClientWidthDefault
  guiImagePreviewClientHeight := guiImagePreviewClientHeightDefault


  marginLeft := 2
  marginRight := 2
  marginTop := 2
  marginBottom := 50

  ;----------------------------- timing variables -----------------------------
  externalAppOpenDelayDefault := 500
  externalAppOpenDelay := externalAppOpenDelayDefault

  commandsDelay := 500
  closeDelay := 15000
  hintDelay := 3000
  copyDelay := 5000
  winRunDelay := 2000

}






















