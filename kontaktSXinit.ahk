; kontaktSXinit.ahk
; Part of kontaktSupporter.ahk

;-------------------------------- readConfig --------------------------------
readConfig(){
  global
  
  if (FileExist(configFile)){
    ; config section:
    fontName := iniReadSave("fontName", "config", "Segoe UI")
    fontsize := iniReadSave("fontsize", "config", 10)
    fontNamePreview := iniReadSave("fontNamePreview", "config", "Consolas")
    fontsizePreview := iniReadSave("fontsizePreview", "config", 11)
    
    baseDirectory := iniReadSave("baseDirectory", "config", A_ScriptDir)
    tableFilePath := iniReadSave("tableFilePath", "config", "ksTabellen")
    selectedTableFile := iniReadSave("selectedTableFile", "config", "ksTabelle.txt")
    bodyFilePath := iniReadSave("bodyFilePath", "config", "ksInhalte")
    attachmentsPath := iniReadSave("attachmentsPath", "config", "ksAnhang")
    textModulsFilePath := iniReadSave("textModulsFilePath", "config", "ksTextBausteine")
    logfileName := iniReadSave("logfileName", "config", "ksGesendet.txt")
    substitutionCharLF := Chr(iniReadSave("substitutionCharLF", "config", 0x00AC))
    separatorChar := Chr(iniReadSave("separatorChar", "config", 0x00A6))
    voiceIsEnabled := iniReadSave("voiceIsEnabled", "config", 0)
    sendDelay := iniReadSave("sendDelay", "config", 5)
    externalAppOpenDelay := iniReadSave("externalAppOpenDelay", "config", externalAppOpenDelayDefault)
    disableScintilla := iniReadSave("disableScintilla", "config", disableScintillaDefault)
    
    externalAppsFile := iniReadSave("externalAppsFile", "external", "ksExternalApps.txt")
    
    dpiScaleValue := iniReadSave("dpiScaleValue", "gui", dpiScaleValueDefault)
    currentLineNumber := iniReadSave("currentLineNumber", "gui", 0)
    
    menuTxtVisible := iniReadSave("menuTxtVisible", "gui", 1)
    guiMainPosX := iniReadSave("guiMainPosX", "gui", guiMainPosXDefault)
    guiMainPosY := iniReadSave("guiMainPosY", "gui", guiMainPosYDefault)
    
    guiMainClientWidth := iniReadSave("guiMainClientWidth", "gui", guiMainClientWidthDefault)
    guiMainClientHeight := iniReadSave("guiMainClientHeight", "gui", guiMainClientHeightDefault)
    
    
    guiPreviewPosX := iniReadSave("guiPreviewPosX", "gui", guiPreviewPosXDefault)
    guiPreviewPosY := iniReadSave("guiPreviewPosY", "gui", guiPreviewPosYDefault)
    
    guiPreviewClientWidth := iniReadSave("guiPreviewClientWidth", "gui", guiPreviewClientWidthDefault)
    guiPreviewClientHeight := iniReadSave("guiPreviewClientHeight", "gui", guiPreviewClientHeightDefault)
    
    
    guiImagePreviewPosX := iniReadSave("guiImagePreviewPosX", "gui", guiImagePreviewPosXDefault)
    guiImagePreviewPosY := iniReadSave("guiImagePreviewPosY", "gui", guiImagePreviewPosYDefault)
    
    guiImagePreviewClientWidth := iniReadSave("guiImagePreviewClientWidth", "gui", guiImagePreviewClientWidthDefault)
    guiImagePreviewClientHeight := iniReadSave("guiImagePreviewClientHeight", "gui", guiImagePreviewClientHeightDefault)
    
  } else {
  FileAppend "
(
[config]
fontName=Segoe UI
fontsize=10
fontNamePreview=Consolas
fontsizePreview=11
baseDirectory=
tableFilePath=ksTabellen
selectedTableFile=ksTabelle.txt
bodyFilePath=ksInhalte
attachmentsPath=ksAnhang
textModulsFilePath=ksTextBausteine
logfileName=ksGesendet.txt
voiceIsEnabled=1
sendDelay=200
externalAppOpenDelay=500
disableScintilla=0
[external]
externalAppsFile=ksExternalApps.txt
[gui]
dpiScale=96
currentLineNumber=2
menuTxtVisible=1
guiMainPosX=0
guiMainPosY=0
guiMainClientWidth=600
guiMainClientHeight=200
guiPreviewPosX=0
guiPreviewPosY=0
guiPreviewClientWidth=800
guiPreviewClientHeight=600
guiImagePreviewPosX=0
guiImagePreviewPosY=0
guiImagePreviewClientWidth=800
guiImagePreviewClientHeight=600
)", configFile, "`n"
  
  }
  
  dpiCorrect := A_ScreenDPI / dpiScaleValue
  
  guiMainPosX := max(guiMainPosX, minPosLeft)
  guiMainPosY := max(guiMainPosY, minPosTop)
  
  guiMainPosX := min(guiMainPosX, A_ScreenWidth - minWidthMargin)
  guiMainPosY := min(guiMainPosY, A_ScreenHeight - minHeightMargin)
  
  
  guiPreviewPosX := max(guiPreviewPosX, minPosLeft)
  guiPreviewPosY := max(guiPreviewPosY, minPosTop)
  
  guiPreviewPosX := min(guiPreviewPosX, A_ScreenWidth - minWidthMargin)
  guiPreviewPosY := min(guiPreviewPosY, A_ScreenHeight - minHeightMargin)
  
  
  guiImagePreviewPosX := max(guiImagePreviewPosX, minPosLeft)
  guiImagePreviewPosY := max(guiImagePreviewPosY, minPosTop)
  
  guiImagePreviewPosX := min(guiImagePreviewPosX, A_ScreenWidth - minWidthMargin)
  guiImagePreviewPosY := min(guiImagePreviewPosY, A_ScreenHeight - minHeightMargin)
}
;-------------------------------- iniReadSave --------------------------------
iniReadSave(name, section, defaultValue){
  global configFile
  
  r := IniRead(configFile, section, name, defaultValue)
  if (r == "" || r == "ERROR")
    r := defaultValue
    
  return r
}
;---------------------------- prepareDirectories ----------------------------
prepareDirectories(){
  global
  local d
  
  if (InStr(FileExist(baseDirectory), "R")){
    msgbox("Ordner `"" baseDirectory "`" (baseDirectory) ist schreibgeschützt!", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
    exit()
  }
  
  if (!FileExist(baseDirectory)){
    msgbox("Ordner `"" baseDirectory "`" (baseDirectory) existiert nicht!", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
    exit()
  }
  
  if (InStr(tableFilePath, ":")){ ; absolute path?
    tableFilePathInUse := tableFilePath
  } else {
    tableFilePathInUse := baseDirectory "\" tableFilePath
  }
  createDir(tableFilePathInUse)
  
  if (InStr(bodyFilePath, ":")){ ; absolute path?
    bodyFilePathInUse := bodyFilePath
  } else {
    bodyFilePathInUse := baseDirectory "\" bodyFilePath
  }
  createDir(bodyFilePathInUse)
  
  if (InStr(protoFilePath, ":")){ ; absolute path?
    protoFilePathInUse := protoFilePath
  } else {
    protoFilePathInUse := baseDirectory "\" protoFilePath
  }
  createDir(protoFilePathInUse)
  
  if (InStr(attachmentsPath, ":")){ ; absolute path?
    attachmentsPathInUse := attachmentsPath
  } else {
    attachmentsPathInUse := baseDirectory "\" attachmentsPath
  }
  createDir(attachmentsPathInUse)
  
  if (InStr(textModulsFilePath, ":")){ ; absolute path?
    textModulsFilePathInUse := textModulsFilePath
  } else {
    textModulsFilePathInUse := baseDirectory "\" textModulsFilePath
  }
  createDir(textModulsFilePathInUse)
}
;--------------------------------- createDir ---------------------------------
createDir(d){
  global
  
  if (!FileExist(d)){
    try {
      DirCreate d 
    } catch Error {
      msgbox("Verzeichnis: `n`"" d "`"`nkann nicht erzeugt werden!`n", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
      exit()
    }
  }
}
;------------------------------- readMainTable -------------------------------
readMainTable(){
  global tableFilePathInUse, selectedTableFile
  local theFile, mainTableContent, code
  
  theFile := ""
  if (InStr(selectedTableFile, ":")){
    theFile := selectedTableFile ; filename is absolute
  } else {
    theFile := tableFilePathInUse "\" selectedTableFile
  }
  
  if (FileExist(theFile)){
    mainTableContent := FileRead(theFile)
  } else {
    code := FileExist(theFile)
    msgbox("Datei `"" theFile "`"`nwurde nicht gefunden!`n`n(readMainTable code: " code ")", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
    exit()
  }
  
  return mainTableContent
}
;--------------------------- parseMainTableContent ---------------------------
parseMainTableContent(mainTableContent){
  global
  local linecontent, msgResult
  
  mTb := []
  mTbLine := []
  currentLine := 0
  Loop Parse, mainTableContent, "`n", "`r" {
    currentLine := A_Index
    linecontent := A_LoopField
    if (StrLen(linecontent)> 1){ ; ignore empty lines
      Loop parse, linecontent, separatorChar {
        if (A_LoopField != "")
          mTbLine.push(A_LoopField)
        else
          mTbLine.push("")
      }
      mTb.push(mTbLine)
      
      ; checkRowIntegrityTableFile
      if (mTbLine.length < mTbCols){
        msg := "Die Datei `"" selectedTableFile "`" ist leider fehlerhaft!"
        msg .= "`nFehler in der Zeile: " currentLine "`n`(" mTbLine.length " statt " mTbCols " Einträge)`n`n"
        msg .= "kontaktSupporter wird beendet!"
        msgbox(msg, "Schwerer Fehler aufgetreten, beende kontaktSupporter!", 48)
        exit()
      }
      mTbLine := []
    }
  }
  
  mTbIndexToNamedVari()
}
;---------------------------- mTbIndexToNamedVari ----------------------------
mTbIndexToNamedVari(){
  global
  
  if (currentLineNumber > 0){
    if (currentLineNumber > mTb.length)
      currentLineNumber := mTb.length
      
    name := mTb[currentLineNumber][1]
    adrText := kontaktSTBinsert(mTb[currentLineNumber][2], 1)
    ccText := kontaktSTBinsert(mTb[currentLineNumber][3], 1)
    bccText := kontaktSTBinsert(mTb[currentLineNumber][4], 1)
    subjectText := mTb[currentLineNumber][5]
    subjectText := (subjectText = "") ? "-" : subjectText ; darf nicht leer sein!
    subjectText := convertStringTo_URI(kontaktSTBinsert(subjectText, 1))
    bodyFileName := autoAppendTxtExt(mTb[currentLineNumber][6])
    attachment := modifyAttachmentsPath(mTb[currentLineNumber][7])
    okField := mTb[currentLineNumber][8]
    memoField := mTb[currentLineNumber][9]
    
    nameRaw := mTb[currentLineNumber][1]
    adrTextRaw := mTb[currentLineNumber][2]
    ccTextRaw := mTb[currentLineNumber][3]
    bccTextRaw := mTb[currentLineNumber][4]
    subjectTextRaw := mTb[currentLineNumber][5]
    subjectTextRaw := (subjectTextRaw = "") ? "-" : subjectTextRaw ; darf nicht leer sein!
    
    bodyFileNameRaw := mTb[currentLineNumber][6]
    attachmentRaw := mTb[currentLineNumber][7]
    okFieldRaw := mTb[currentLineNumber][8]
    memoFieldRaw := mTb[currentLineNumber][9]
  }
}
;----------------------------- readExternalApps -----------------------------
readExternalApps(){
  global 
  local theFile, linecontent, externalAppsLine
  
  theFile := ""
  if (InStr(externalAppsFile, ":")){
    theFile := externalAppsFile ; filename is absolute
  } else {
    theFile := baseDirectory "\" externalAppsFile
  }
  
  if (FileExist(theFile)){
    externalAppsContent := FileRead(theFile)
    
    externalApps := []
    externalAppsLine := []
    currentLine := 0
    Loop Parse, externalAppsContent, "`n", "`r" {
      currentLine := A_Index
      linecontent := A_LoopField
      if (StrLen(linecontent) > 0){
        if (SubStr(linecontent,1,1) != ";" && SubStr(linecontent,1,1) != "#" && SubStr(linecontent,1,2) != "//"){
          Loop parse, linecontent, separatorChar, "`r" {
            externalAppsLine.push(A_LoopField)
          }
          ; checkRowIntegrityExternalAppsFile
          if (externalAppsLine.length < 2){
            msgbox("Die Datei `"" externalAppsFile "`"`nist leider fehlerhaft!`n`n(Fehler in der Zeile: " currentLine " )!`n`n(checkRowIntegrityTableFile)", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
          }
          
          externalApps.push(externalAppsLine)
          externalAppsLine := []
        }
      } else {
        externalAppsLine.push("")
        externalApps.push(externalAppsLine)
        externalAppsLine := []
      }
    }
  } else {
    msgbox("Datei `"" theFile "`" nicht gefunden (readExternalApps)!", "Schwerer Fehler aufgetreten, App wird beendet!", "Icon!")
    exit()
  }
}





