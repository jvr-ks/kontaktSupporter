/*
 *********************************************************************************
 * 
 * kontaktSupporter.ahk
 * 
 * uses UTF-8
 * 
 * Version: look at appVersion := below
 * 
 * Copyright (c) 2024 jvr.de. All rights reserved.
 *
 *
 *********************************************************************************
*/
/*
 *********************************************************************************
 * 
 * GNU GENERAL PUBLIC LICENSE
 * 
 * A copy is included in the file "license.txt"
 *
  *********************************************************************************
*/
/*
 *********************************************************************************
 * 
 * Function to send is: "LV1_DoubleClick" -> "sendSelected" -> "sendTxtOnly"
 * 
 * 
 *
  *********************************************************************************
*/


#Requires AutoHotkey v2

#Warn
#SingleInstance force

#DllLoad "*i C:\Windows\System32\GdiPlus.dll"

Fileencoding "UTF-8-RAW"

#Include kontaktSXVariables.ahk
#Include kontaktSXinsert.ahk
#Include kontaktSXrowEdit.ahk
#Include kontaktSXinit.ahk
#Include kontaktSXmenu.ahk
#Include kontaktSXhelper.ahk
#Include kontaktSXScintilla2.ahk
#Include kontaktSXguiPreview.ahk
#Include lib\Gdip_All2.ahk

SetTitleMatchMode "2"
DetectHiddenWindows false
SendMode "Input" 

InstallKeybdHook true

;-------------------------------- read cmdline param --------------------------------
hasParams := A_Args.Length
starthidden := 0

if (hasParams != 0){
  Loop hasParams
  {
    if(A_Args[A_index] == "remove"){
      ExitApp 0
    }

    if(A_Args[A_index] == "hidewindow"){
      starthidden := 1
    }
  }
}

baseDirectory := A_ScriptDir

appName := "kontaktSupporter"
appnameLower := "kontaktsupporter"
extension := ".exe"
appVersion := "0.066"
app := appName . " " . appVersion

appTitle := appName " " "v" appVersion

configFileHidden := "_"  appnameLower ".ini"

configFile := appnameLower ".ini"
  
if (FileExist(configFileHidden))
  configFile := configFileHidden

initGlobalVariables()

;----------------------------------- init -----------------------------------
readConfig()
prepareDirectories()
parseMainTableContent(readMainTable())
readExternalApps()

setMenuTxt()
mainGui()
previewGui()
imagePreviewGui()

HotIfWinActive "ahk_class AutoHotkeyGUI"
hotkey("F1", quickHelp, "On")

;-------------------------------- Start gdi+ --------------------------------; 
pToken := Gdip_Startup()
if (!pToken) {
  MsgBox "Gdiplus konnte nicht gestartet werden, Bilderdarstellung ist eingeschränkt!"
}

return

;---------------------------------- mainGui ----------------------------------
mainGui(){
  global
  local LV1width, LV1height, msg, entry, missingEntries, theExternalName
  
  guiMain := Gui("+OwnDialogs +LastFound MaximizeBox +Resize", appTitle)
  guiMain.SetFont("S" . fontsize, fontName)
  
  
  TableMenu := Menu()
  TableMenu.Add("Neuer Tabellen-Eintrag (leer)", newEntryTableFile)
  TableMenu.Add("Neuer Tabellen-Eintrag (Kopie des gewählten)", duplicateEntryTableFile)
  TableMenu.Add("Gewählten Tabellen-Eintrag löschen", deleteEntryTableFile)
  TableMenu.Add("Tabelle bearbeiten (Dateitext)", editselectedTableFile)
  TableMenu.Add("Tabelle wählen", selectTableFile)
  TableMenu.Add("Ordner (Tabelle) öffnen", openTableFolder)
  TableMenu.Add("Trennzeichen in die Zwischenablage laden", separatorCharToClipBoard)
  
  ConfigMenu := Menu()
  ConfigMenu.Add("Gesendet (Aufstellung)`"" logfileName "`"  ansehen", editLogFile)
  ConfigMenu.Add("")
  ConfigMenu.Add("Konfiguration`"" configFile "`" bearbeiten", editConfig)
  ConfigMenu.Add("Ordner (App) öffnen", openBaseDirectory)
  ConfigMenu.Add("")
  ConfigMenu.Add("Externe Apps Liste `"" externalAppsFile "`" erstellen/bearbeiten", editExternalApps)
  
  AttachmentMenuGlobal := Menu()
  AttachmentMenuGlobal := attachmentMenuBuildGlobal(AttachmentMenuGlobal)
  
  mainMenu := MenuBar()
  mainMenu.Add("📖" menuItem1, rowEdit)
  mainMenu.Add("♒" menuItem2, rowPreview)
  mainMenu.Add("📂" menuItem3, TableMenu)
  mainMenu.Add("🔗" menuItem4, AttachmentMenuGlobal)
  mainMenu.Add("⚙" menuItem5, ConfigMenu)
  mainMenu.Add("🛈" menuItem6, openHelp)
  mainMenu.Add("⟳" menuItem7, refresh)
  mainMenu.Add("" menuItem8, toggleMenuTxt)
  mainMenu.Add("🗙" menuItem9, exit)
  
  guiMain.MenuBar := mainMenu
  
  LV1 := guiMain.AddListView("r30 -Multi", ["", "Name", "Adresse", "Cc", "Bcc",
  "Subject", "Inhalt (Dateiname)", "Anhang (Pfad)", "Ok", "Memo"])
  
  msg := " Hilfe: Einfachklick zum Markieren, Doppelklick zum Senden ... weiteres F1-Taste drücken! "
  msg .= "[Tabelle: " selectedTableFile "] "
  msg .= "[ConfigFile: " configFile "] "
  
  statusBar := guiMain.AddStatusBar(, msg)
  
  guiMain.show("x" guiMainPosX " y" guiMainPosY " w" guiMainClientWidth " h" guiMainClientHeight)
  mainIsVisible := 1

  LV1.OnEvent("Click", LV1_Click)
  LV1.OnEvent("DoubleClick", LV1_DoubleClick)
  
  LV1.Opt("-Redraw")
  loop mTb.length {
    LV1.Add(, format("{:03}", A_Index), mTb[A_Index][1], mTb[A_Index][2],
    mTb[A_Index][3], mTb[A_Index][4], mTb[A_Index][5], shortenDisplayed(mTb[A_Index][6]),
    shortenDisplayed(mTb[A_Index][7]), mTb[A_Index][8], shortenDisplayed(mTb[A_Index][9]))
  }
  LV1.Opt("+Redraw")
  
  if (selectedIsValid(currentLineNumber))
    LV1.Modify(currentLineNumber, "+Select") ; select actual row if valid
  
  guiMain.OnEvent("Size", guiMain_Size , 1)
  guiMain.OnEvent("Close", exit)
  OnMessage(0x03, moveEventSwitch)
   
  if (guiMainClientWidth < guiMainClientWidthDefault || guiMainClientHeight < guiMainClientHeightDefault || guiMainClientWidth > coordsScreenToApp(A_ScreenWidth) || guiMainClientHeight > coordsScreenToApp(A_ScreenHeight)){

    guiMain.show("x" 0 " y" 0 " w" guiMainClientWidthDefault " h" guiMainClientHeightDefault)

    ; trigger a "Size" event to resize LV1
    guiMainHwnd := guiMain.hwnd
    WinMinimize "A"
    WinRestore "ahk_id " guiMainHwnd
  } else {
    LV1width := (guiMainClientWidth - LV1MarginLeft - LV1MarginRight)
    LV1height := (guiMainClientHeight - LV1MarginTop - LV1MarginBottom)
    
    LV1.Move(LV1MarginLeft, LV1MarginTop, LV1width, LV1height)
  }
  
  LV1.Opt("-Redraw")
  ; Breite Felder begrenzen (in Pixel!)
  LV1.ModifyCol()
  LV1.ModifyCol(1, "30 Integer") ; No
  LV1.ModifyCol(2, "AutoHdr") ; Name
  LV1.ModifyCol(3, "AutoHdr") ; Adresse
  LV1.ModifyCol(4, "AutoHdr") ; Cc
  LV1.ModifyCol(5, "AutoHdr") ; Bcc
  LV1.ModifyCol(6, "AutoHdr") ; Subject
  LV1.ModifyCol(7, "AutoHdr") ; Inhalt (Dateiname)
  LV1.ModifyCol(8, "AutoHdr") ; Anhang Pfad
  LV1.ModifyCol(9, "AutoHdr") ; Ok
  LV1.ModifyCol(10, 300) ; Memofeld
  LV1.Opt("+Redraw")
}
;------------------------------ moveEventSwitch ------------------------------
moveEventSwitch(p1, p2, p3, p4, *){
  local h1, h2, h3
  
  h1 := 0, h2 := 0, h3 := 0
  
  if (IsSet(guiMain)){
    h1 := guiMain.hwnd
  }
    
  if (IsSet(guiPreview)){
    h2 := guiPreview.hwnd
  }
  
  if (IsSet(guiImagePreview)){
    h3 := guiImagePreview.hwnd
  }
  
  Switch  p4
  {
    Case h1:
      guiMainMove()
    
    Case h2:
      guiPreviewMove()
      
    Case h3:
      guiImagePreviewMove()
  
  }
}
;------------------------- separatorCharToClipBoard -------------------------
separatorCharToClipBoard(*){
  global
  
  A_Clipboard := separatorChar
}
;-------------------------------- mouseMoved --------------------------------
WM_MOUSEMOVED(*){
  global 
   
  OnMessage 0x200, WM_MOUSEMOVED, 0
  Reload
}
;------------------------------- guiMain_Size -------------------------------
guiMain_Size(thisGui, MinMax, guiMainClientWidth, guiMainClientHeight) {
  global LV1MarginTop, LV1MarginBottom, LV1MarginLeft, LV1MarginRight
  local LV1width, LV1height
  
  if (MinMax = -1)
      return
      
  IniWrite guiMainClientWidth, configFile, "gui", "guiMainClientWidth"
  IniWrite guiMainClientHeight, configFile, "gui", "guiMainClientHeight"
  
  LV1width := (guiMainClientWidth - LV1MarginLeft - LV1MarginRight)
  LV1height := (guiMainClientHeight - LV1MarginTop - LV1MarginBottom)
  
  LV1.Move(LV1MarginLeft, LV1MarginTop, LV1width, LV1height)
}
;-------------------------------- guiMainMove --------------------------------
guiMainMove(){
  global
  
  if (IsSet(guiMain)){
    guiMain.GetPos(&guiMainPosX, &guiMainPosY)
    
    guiMainPosX := max(guiMainPosX, minPosTop)
    guiMainPosY := max(guiMainPosY, minPosLeft)
      
    IniWrite guiMainPosX, configFile, "gui", "guiMainPosX"
    IniWrite guiMainPosY, configFile, "gui", "guiMainPosY"
  }
}
;--------------------------------- LV1_Click ---------------------------------
LV1_Click(LV1, linenumber){
  global
  
  currentLineNumber := linenumber
  IniWrite currentLineNumber, configFile, "gui", "currentLineNumber"
  
  mTbIndexToNamedVari()
  
  if (IsSet(guiImagePreview))
    guiImagePreview.Hide()
  
  if (GetKeyState("Shift")){
    imagePreview()
    
    return
  }

  if (displayPreview){
    guiPreview_Close()
    guiImagePreview_Close()
    rowPreview()
  }
  
  if (rowEditIsVisible){
    guiRowEditClose()
    rowEdit()
  }
  
}
;------------------------------ LV1_DoubleClick ------------------------------
LV1_DoubleClick(LV1, linenumber){
  global   
  currentLineNumber := linenumber
  IniWrite currentLineNumber, configFile, "gui", "currentLineNumber"
    
  mTbIndexToNamedVari()

  sendSelected(0)
  guiMain.show
}
;---------------------------------- LV1Save ----------------------------------
LV1Save(){
  global
  local lineNumber
  
  s := ""
  bIndex := 0
  loop mTb.length {
    lineNumber := A_Index
    loop (mTbCols) {
      s .= removeEmptyChar(mTb[lineNumber][A_Index]) separatorChar
    }
    s := SubStr(s, 1, -1)
    s .= "`n"
  }
  
  mainTableContent := s
  writeMainTable(s)
  
  mTbIndexToNamedVari()
}
;-------------------------------- LV1_Update --------------------------------
LV1Update(){
  global
  local lineNumber
  
  loop mTb.length {
    lineNumber := A_Index
    loop mTbCols {
      col := "Col" (A_Index + 1)
      LV1.Modify(lineNumber, col, mTb[lineNumber][A_Index])
    }
  }
  if (selectedIsValid(currentLineNumber))
    LV1.Modify(currentLineNumber, "+Select") ; select actual row
  
}
;--------------------------------- quickHelp ---------------------------------
quickHelp(*){
  global
  local msg, width, height, opt, wb
  
  if (!quickIsHelpVisible){
    guiShortHelp := Gui("+OwnDialogs +Resize +parent" guiMain.Hwnd, "Kurz-Hilfe")
    
    width := round(A_ScreenWidth * 0.9)
    height := round(A_ScreenHeight * 0.9)
    guiShortHelp.Add("Edit", "x2 y2 r0 w0 h0", "") ; Focus dummy
    wb := guiShortHelp.Add("ActiveX", "x2 y2 w" width " h" height " vWB", "about:<!DOCTYPE html><meta http-equiv=`"X-UA-Compatible`" content=`"IE=edge`"").Value
    ;ComObjConnect(wb, wb_events)
    doc := wb.document
    doc.documentElement.style.overflow := "scroll"
    quickHelpContent := FileRead(A_ScriptDir "\" "kontaktSXQHelp.html")
    doc.write(quickHelpContent)
    
    guiShortHelp.Opt("+Owner" guiMain.Hwnd)
    
    guiShortHelp.show("center autosize")
    
    guiShortHelp.OnEvent("Close", guiShortHelp_Close)
  } else {
    guiShortHelp.destroy
  }
  
  quickIsHelpVisible := !quickIsHelpVisible
}
;---------------------------- guiShortHelp_Close ----------------------------
guiShortHelp_Close(*){
  global quickIsHelpVisible
  
  quickIsHelpVisible := 0
}
;----------------------------- newEntryTableFile -----------------------------
newEntryTableFile(*){
  global
  local theNewFile, s, theFile
  
  ; create the new body-file "neu.txt"
  theNewFile := bodyFilePathInUse "\neu.txt"
  if (!FileExist(theNewFile))
    FileAppend("noch leer ...", theNewFile, "`n")
  
  ; append a new entry
  s := "NEU!" separatorChar     ;name
  s .= "xy@xy.de" separatorChar ; adrText
  s .= "" separatorChar         ; ccText
  s .= "" separatorChar         ; bccText
  s .= "-" separatorChar        ; subjectText
  s .= "leer" separatorChar     ; bodyFileName
  s .= "" separatorChar         ; attachment
  s .= "" separatorChar         ; okField
  s .= ""                       ; memoField
  s .= "`n"
  
  theFile := ""
  if (InStr(selectedTableFile, ":")){
    theFile := selectedTableFile ; filename is absolute
  } else {
    theFile := tableFilePathInUse "\" selectedTableFile
  }
  FileAppend(s, theFile, "`n")
  sleep 200
  
 ; readback
  parseMainTableContent(readMainTable())
  LV1AddEntry()
}
;-------------------------- duplicateEntryTableFile --------------------------
duplicateEntryTableFile(*){
  global
  local s, theFile
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    ; append the duplicated entry
    s := nameRaw separatorChar         ;name
    s .= adrTextRaw separatorChar      ; adrText
    s .= ccTextRaw separatorChar       ; ccText
    s .= bccTextRaw separatorChar      ; bccText
    s .= subjectTextRaw separatorChar  ; subjectText
    s .= bodyFileNameRaw separatorChar ; bodyFileName
    s .= attachmentRaw separatorChar   ; attachment
    s .= okFieldRaw separatorChar      ; okField
    s .= memoFieldRaw                  ; memoField
    s .= "`n"
    
    theFile := ""
    if (InStr(selectedTableFile, ":")){
      theFile := selectedTableFile ; filename is absolute
    } else {
      theFile := tableFilePathInUse "\" selectedTableFile
    }
    FileAppend(s, theFile, "`n")
    sleep 200
  
   ; readback
    parseMainTableContent(readMainTable())
    
    LV1AddEntry()
  }
}
;--------------------------- deleteEntryTableFile ---------------------------
deleteEntryTableFile(*){
  global
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    mTb.RemoveAt(currentLineNumber)
    LV1Save()
    sleep 500
    reload
  }
}
;-------------------------------- LV1AddEntry --------------------------------
LV1AddEntry(){
  global
  
  lastEntry := mTb.Length
  LV1.Add(, format("{:03}", lastEntry), mTb[lastEntry][1], mTb[lastEntry][2],
  mTb[lastEntry][3], mTb[lastEntry][4], mTb[lastEntry][5], shortenDisplayed(mTb[lastEntry][6]),
  shortenDisplayed(mTb[lastEntry][7]), mTb[lastEntry][8], shortenDisplayed(mTb[lastEntry][9]))
}
;------------------------------ removeEmptyChar ------------------------------
; only used to remove char from old files ...
removeEmptyChar(s){
  s := (s != "-") ? s : ""
  
  return s
}
;----------------------------- encodeNotAllowed -----------------------------
encodeNotAllowed(s){
  global substitutionCharLF
  
  s := StrReplace(s, "`n", substitutionCharLF)
  
  return s
}
;----------------------------- decodeNotAllowed -----------------------------
decodeNotAllowed(s){
  global substitutionCharLF
  
  s := StrReplace(s, substitutionCharLF, "`n")
  
  return s
}
;----------------------------- shortenDisplayed -----------------------------
shortenDisplayed(s){
  if (StrLen(s) > 35){
    s := SubStr(s, 1, 10) " ... " SubStr(s, -20,20)
  }
  return s
}
;------------------------------ openHelp ------------------------------
openHelp(*){
  global baseDirectory

  toRun := baseDirectory "\kontaktsupporter.html"
  
  if (FileExist(toRun)){
    RunWait toRun
  } else {
    msgbox("Datei `"" toRun "`" nicht gefunden (openHelp)!", "Fehler aufgetreten!", "Icon!")
  }
}
;------------------------------ selectedIsValid ------------------------------
selectedIsValid(selected){
  global mTb
  local valid, msgResult
  
  valid := ((selected < 1) || (selected > mTb.length)) ? 0 : 1

  return valid
}
;------------------------------- sendSelected -------------------------------
sendSelected(preview := 0){
  global
  local clipboardSave, data, msg, msgP, param
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
  
    if (GetKeyState("Shift")){
      imagePreview()
      
      return
    }
    
    clipboardSave := ClipboardAll()

    if (!preview)
      guiMain.hide
      
    bodyText := ""
    openLink := ""
    
    switchTags()
    combineOpenLink()
      
    if ((SubStr(bodyFileName, -4) == ".txt") || bodyFileName == ""){
      ; *.txt section
      
      readTxtFile()
      
      if (preview){
        rowPreviewShow()
        
        rowPreviewSetText(bodyText)
        
        generatePreviewParam()
        
        A_Clipboard := clipboardSave 
        clipboardSave := ""
    
        return
      } else {
        currentRunningApps := getRunningAppsAsString()
        
        eMailAppId := detectNewRunningAppLoop(openLink, currentRunningApps)
        
        if (eMailAppId = 0){
          msgbox("Keine neu gestartete eMail-App gefunden ()!", "Schwerer Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
          exit()
        }
        
        sendTxtOnly(bodyText)
      }
    } else {
      ; other than *.txt section: try to open the file with the app assigned to the file extension
      
      if (preview){
        rowPreviewShow()
        rowPreviewSetText("Inhalt ist kein Textformat, bitte auf `"🖼 Anhang Vorschau`" klicken!")
        
        generatePreviewParam()
        
        A_Clipboard := clipboardSave 
        clipboardSave := ""
    
        return
      } else {
        currentRunningApps := getRunningAppsAsString()
        
        externalAppName := detectNewRunningAppLoop(getBodyFile(), currentRunningApps)
        
       if (externalAppName == ""){
          msgbox("Keine neu gestartete App gefunden!", "Schwerer Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
          exit()
        }
        
        WinActivate(externalAppName)
        
        SendInput("^a")
        sleep commandsDelay
        SendInput("^c")
        sleep copyDelay

        if (WinExist(externalAppName)){
          WinClose externalAppName
          ;ProcessClose PID oder Name
          sleep closeDelay
        }
        
        currentRunningApps := getRunningAppsAsString()
        
        eMailAppId := detectNewRunningAppLoop(openLink, currentRunningApps)
        
        if (eMailAppId = 0){
          msgbox("Keine neu gestartete eMail-App gefunden!", "Schwerer Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
          exit()
        }
        
        WinActivate("ahk_id " eMailAppId)
        sleep sendDelay
        SendInput("^v")
        sleep sendDelay
        SendInput("^{Home}")
        sleep sendDelay
        
        msgP := "Inhalt wurde eingefügt und sollte jetzt manuell überprüft/bearbeitet werden!`n`n"
        msg := leaveMessage(msgP,)
      
        showHintColored(msg, 0,,,,,"y10 xcenter")
        WinActivate("ahk_id " eMailAppId)
        
        msgResult := msgbox("Die eMail als gesendet eintragen?", "eMail Bearbeitung fertig und gesendet?" , 0x21)

        if (msgResult = "OK"){
          dateData := FormatTime(, "yyyy/MM/dd ") separatorChar FormatTime("T12", "Time") separatorChar trim(name) separatorChar trim(adrText) separatorChar trim(ccText) separatorChar trim(bccText) separatorChar trim(subjectText) separatorChar trim(bodyFileName) separatorChar attachment "`n"

          dateData := RegExReplace(dateData,"%.{2}"," ")
          if (StrLen(dateData) > 70)
            dateData .= "`n"
          
          FileAppend dateData, logfileName, "`n"
        } else {
          showHintColored("Es erfolgte kein Eintrag in der Datei `"gesendet.txt`" !", hintDelay)
        }
  
        hintColored.Destroy()
        
      }
      FileDelete "ClipBoard*.txt"
      
      if (WinExist("ahk_id " eMailAppId)){
        WinClose "ahk_id " eMailAppId
        sleep closeDelay
      }
      
      if (WinExist("ahk_id " eMailAppId))
        msgbox "Konnte die eMail-App nicht schließen,`nbitte die eMail-App manuell beenden!"
    }
    A_Clipboard := clipboardSave 
    clipboardSave := ""
    
    if (preview)
      guiMain.show
  }
}
;------------------------------- leaveMessage -------------------------------
leaveMessage(prepend := "", append := ""){
  global
  local msg

  msg := prepend
 
  if (attachment != ""){
    A_Clipboard := modifyAttachmentsPath(attachment)
    msg .= "`nDer Pfad des Anhangs:`n`n" . attachment . "`n`nist in der Zwischenablage, bitte den Anhang jetzt manuell einfügen.`n`n"
  }
  msg .= "Danach bitte die eMail manuell senden oder abbrechen und die eMail app schließen!`n`n"
  msg .= append
  
  return msg
}
;-------------------------------- readTxtFile --------------------------------
readTxtFile(){
  global bodyFileName, bodyText
  local theFile
  
  bodyText := ""
  theFile := ""
  if (bodyFileName != ""){
    theFile := calculateFilePath(bodyFileName)
    if (FileExist(theFile)){
      bodyText := kontaktSTBinsert(FileRead(theFile))
    } else {
      msgbox("WARNUNG! Inhalts-Datei-Feld ist leer oder Datei wurde nicht gefunden!`n`n(readTxtFile)")
    }
  }
}
;----------------------------- calculateFilePath -----------------------------
calculateFilePath(f){
  global bodyFilePathInUse
  local r
  
  r := ""
  if (InStr(f, ":")){ ; filename includes directory?
    r := f
  } else {
    r := bodyFilePathInUse "\" f
  }
  
  return r
}
;------------------------------ writeMainTable ------------------------------
writeMainTable(mainTableContent){
  global tableFilePathInUse, selectedTableFile
  local theFile, dateData, tmpFile
  
  theFile := ""
  if (InStr(selectedTableFile, ":")){
    theFile := selectedTableFile ; filename is absolute
  } else {
    theFile := tableFilePathInUse "\" selectedTableFile
  }
  
  tmpFile := tableFilePathInUse "\" StrReplace(selectedTableFile,".txt","") "_previous.txt"
   
  if (FileExist(tmpFile)){
    FileDelete(tmpFile)
    sleep 200
  }
  
  FileAppend(mainTableContent, tmpFile, "`n")
  sleep 500
    
  if (FileExist(theFile)){
    FileDelete(theFile)
    sleep 200
  }
  
  FileAppend(mainTableContent, theFile, "`n")

  return
}
;------------------------------ openTableFolder ------------------------------
openTableFolder(*){
  run tableFilePathInUse
}
;------------------------------ getBodyFilePath ------------------------------
getBodyFile(){
  global bodyFilePathInUse, bodyFileName
  local theFile
  
  theFile := calculateFilePath(bodyFileName)
  theFile := autoAppendTxtExt(theFile)
  if (!FileExist(theFile)){
    msgbox("Datei`n" theFile "`nnicht gefunden (getBodyFile)!", "Schwerer Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
    exit()
  }
  
  return theFile
}
;------------------------------ combineOpenLink ------------------------------
combineOpenLink(){
  global
  
  openLink := adrTag . adrText
  openLink := openLink . ccTag . ccText
  openLink := openLink . bccTag . bccText
  openLink := openLink . subjectTag . subjectText
}
;-------------------------------- switchTags --------------------------------
switchTags(){
  global

  ccTag := (ccText != "") ? "?cc=" : ""
  bccTag := (bccText != "") ? "?bcc=" : ""
  subjectTag := (subjectText != "") ? "?subject=" : ""
  
  if (ccText != ""){
    bccTag := StrReplace(bccTag, "?", "&")
    subjectTag := StrReplace(subjectTag,"?","&")
  }
   
  if (bccText != ""){
    subjectTag := StrReplace(subjectTag,"?","&")
  }
}
;-------------------------------- sendTxtOnly --------------------------------
sendTxtOnly(textContent){
  global 
  local clipboardSave, Errorlevel, dateData, msg
  
  clipboardSave := ClipboardAll()
  
  WinActivate("ahk_id " eMailAppId)
  
  if (textContent != ""){
    A_Clipboard := textContent
    Errorlevel := ClipWait(5,0) ; wait 5 seconds max.
    if (Errorlevel){
      SendInput("^v")
      sleep(sendDelay)
     } else {
      msgbox("Clipboard funktioniert nicht, beende App!", "Schwerer Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
      A_Clipboard := clipboardSave 
      clipboardSave := ""
      exit()
    }
    
    ; select all
    SendInput("^a")
    sleep sendDelay
    
    msg := leaveMessage("Alles bleibt markiert, um Text und Schrifteinstellungen anpassen zu können!`n")
    
    showHintColored(msg, 0,,,,,"y10 xcenter")
  } else {
    msg := leaveMessage("EMail Body-Feld ist leer!`n`n")
    showHintColored(msg, 0,,,,,"y10 xcenter")
  }
  
  WinActivate("ahk_id " eMailAppId)
  
  msgResult := msgbox("Die eMail als gesendet eintragen?", "eMail Bearbeitung fertig und gesendet?" , 0x21)

  if (msgResult = "OK"){
    dateData := FormatTime(, "yyyy/MM/dd ") separatorChar FormatTime("T12", "Time") separatorChar trim(name) separatorChar trim(adrText) separatorChar trim(ccText) separatorChar trim(bccText) separatorChar trim(subjectText) separatorChar trim(bodyFileName) separatorChar attachment "`n"

    dateData := RegExReplace(dateData,"%.{2}"," ")
    if (StrLen(dateData) > 70)
      dateData .= "`n"
    
    FileAppend dateData, logfileName, "`n"
  } else {
    showHintColored("Es erfolgte kein Eintrag in der Datei `"gesendet.txt`" !", hintDelay)
  }

  SendInput("^{Home}")
  hintColored.Destroy()
 
  if (WinExist("ahk_id " eMailAppId)){
    WinClose "ahk_id " eMailAppId
    sleep closeDelay
  }
  
  if (WinExist("ahk_id " eMailAppId))
    msgbox "Konnte die eMail-App nicht schließen, bitte manuell beenden!"
        
  FileDelete "ClipBoard*.txt"
      
  A_Clipboard := clipboardSave 
  clipboardSave := ""
}

;------------------------------ selectTableFile ------------------------------
selectTableFile(*){
  global selectedTableFile, tableFilePathInUse
  local listPath, newselectedTableFile
  
  filesCount := 0
  Loop Files, tableFilePathInUse "\*.*" {
    if A_LoopFileAttrib ~= "[HRS]"
    continue 
    filesCount += 1
  }

  if (!filesCount){
    FileAppend "Name|To eMail|Cc|Bcc|Subject|Data FileName|Anhang FileName", tableFilePathInUse . "\neueTabelle.txt", "`n"
  }
    
  newselectedTableFile := FileSelect("1", tableFilePathInUse, "Liste wählen")
  
  if (newselectedTableFile != ""){
    selectedTableFile := pathToRelativ(newselectedTableFile, tableFilePathInUse)
    IniWrite selectedTableFile, configFile, "config", "selectedTableFile"
    
    Reload
  }
}
;---------------------------- convertStringTo_URI ----------------------------
convertStringTo_URI(s){
; based on: https://rosettacode.org/wiki/UTF-8_encode_and_decode#AutoHotkey

  UTFCode8 := ""
  result := ""
  
  sBasic := StrReplace(s, "%", "%25")
  sBasic := StrReplace(sBasic, chr(09), "%09")
  
  Loop Parse sBasic, "`n", "`r"
  {
    hex := format("{1:#6.6X}", Ord(A_LoopField))
    Bytes :=  hex>=0x10000 ? 4 : hex>=0x0800 ? 3 : hex>=0x0080 ? 2 : hex>=0x0001 ? 1 : 0
    Prefix := [0, 0xC0, 0xE0, 0xF0]
    if (hex>=0x0080){
      UTFCode8 := ""
      Loop Bytes {
        if (A_Index < Bytes)
          UTFCode8 := "%" Format("{:X}", (hex&0x3F) + 0x80) . UTFCode8    ; 3F=00111111, 80=10000000
        else
          UTFCode8 := "%" Format("{:X}", hex + Prefix[Bytes]) . UTFCode8  ; C0=11000000, E0=11100000, F0=11110000
        hex := hex>>6
      }
    } else {
      UTFCode8 := A_LoopField
    }
    result .= UTFCode8
  }
  
  result := StrReplace(result, " ", "%20")
  result := StrReplace(result, "`n", "%0A")
  result := StrReplace(result, "`r`n", "%0A")
  
  result := StrReplace(result, "?", "%3F")
  result := StrReplace(result, "&", "%26")
  
  return result
}
;---------------------------- wmMouseMoveDelay(){ ----------------------------
wmMouseMoveDelay(){
  OnMessage 0x200, WM_MOUSEMOVED, 1 ; enable WM_MOUSEMOVE = 0x200
}
;--------------------------- editselectedTableFile ---------------------------
editselectedTableFile(*){
  global baseDirectory, selectedTableFile, tableFilePath
  local theFile, thePath, toRun 
  
  toRun := ""
  theFile := selectedTableFile
  if (InStr(theFile, ":")){ ; filename includes directory?
    toRun := theFile
  } else {
    toRun := baseDirectory "\" tableFilePath "\" theFile
  }
  if (!FileExist(toRun)){
    FileAppend "", toRun, "`n"
  }
  
  if (FileExist(toRun)){
    ; triggers a reload on the next mousemove event!
    SetTimer wmMouseMoveDelay, -2000
    RunWait toRun
  } else {
    msgbox("Datei `"" toRun "`" nicht gefunden (editselectedTableFile)!", "Fehler aufgetreten!", "Icon!")
  }
}
;-------------------------------- editLogFile --------------------------------
editLogFile(*){
  global baseDirectory, logfileName
  local theFile, toRun 
  
  toRun := ""
  theFile := logfileName
  if (InStr(theFile, ":")){ ; filename includes directory?
    toRun := theFile
  } else {
    toRun := baseDirectory "\" theFile
  }
  if (!FileExist(toRun)){
    FileAppend "", toRun, "`n"
  }
  
  if (FileExist(toRun)){
    RunWait toRun
  } else {
    msgbox("Datei `"" toRun "`" nicht gefunden!", "Fehler aufgetreten!", "Icon!")
  }
}
;-------------------------------- editConfig --------------------------------
editConfig(*){
  global baseDirectory, configFile
  local toRun 
  
  toRun := configFile

  if (FileExist(toRun)){
    SetTimer wmMouseMoveDelay, -2000
    RunWait toRun
  } else {
    msgbox("Datei `"" toRun "`" nicht gefunden!", "Fehler aufgetreten!", "Icon!")
  }
}
;----------------------------- openBaseDirectory -----------------------------
openBaseDirectory(*){
  global baseDirectory, configFile
  
  run A_ComSpec " /c start " baseDirectory
}
;----------------------------- editExternalApps -----------------------------
editExternalApps(*){
  global baseDirectory, externalAppsFile
  local toRun 
  
  toRun := baseDirectory "\" externalAppsFile
  if (!FileExist(toRun)){
    FileAppend "", toRun, "`n"
  }
  
  if (FileExist(toRun)){
    SetTimer wmMouseMoveDelay, -2000
    RunWait toRun
  } else {
    msgbox("Datei `"" toRun "`" nicht gefunden!", "Fehler aufgetreten!", "Icon!")
  }
}
;------------------------------- openExternal -------------------------------
openExternal(externalAppsName, toRun, *){
  global
  local d, found, extractedExternalPath
  
  d := ""
  found := 0

    
  if (!FileExist(toRun)){
    showHintColored("Externe App existiert nicht (Datei `"" toRun "`"), verwende default!", 5000)
    run A_ComSpec " /c start " d 
    
    return
  }
  
  ; explorer:
  if (InStr(externalAppsName, "explorer")){
    found := RegExMatch(externalAppsName, "\[(.+?)\]", &extractedExternalPath, 1)
    if (!found){
      ; attachment can be a folder, or a file
      d := modifyAttachmentsPath(attachment)
      msgbox aquot toRun aqbq d aquot
      run aquot toRun aqbq d aquot
    } else {
      
      toRun := RegExReplace(toRun, "\[(.+?)\]","") ; remove [...]
      run aquot toRun aqbq extractedExternalPath[1] aquot
    }
    return
  }
  
  ; Dopus:
  if (InStr(externalAppsName, "dopus")){
    found := RegExMatch(externalAppsName, "\[(.+?)\]", &extractedExternalPath, 1)
    if (!found){
      d := modifyAttachmentsPath(attachment)
      run aquot toRun aqb "/acmd go" abq d aquot
    } else {
      toRun := RegExReplace(toRun, "\[(.+?)\]","") ; remove [...]
      d := modifyAttachmentsPath(attachment)
      run aquot toRun aqb "/acmd go" abq extractedExternalPath[1] aqb "OPENINLEFT DUALPATH" abq d aquot
    }
    return
  }
  ; Freecommander or multicommander:
  if (InStr(externalAppsName, "freecommander") || InStr(externalAppsName, "multicommander")){
    found := RegExMatch(externalAppsName, "\[(.+?)\]", &extractedExternalPath, 1)
    if (!found){
      d := modifyAttachmentsPath(attachment)
      
      run aquot toRun aqb "/OPEN" abq d aquot
    } else {
      toRun := RegExReplace(toRun, "\[(.+?)\]","") ; remove [...]
      d := modifyAttachmentsPath(attachment)
      
      run aquot toRun aqb "/L=" aquot extractedExternalPath[1] aqb "/R=" aquot d aquot
    }
    return
  }
  ; ACDSee:
  if (InStr(externalAppsName, "ACDSee")){
    found := RegExMatch(externalAppsName, "\[(.+?)\]", &extractedExternalPath, 1)
    if (!found){
      d := modifyAttachmentsPath(attachment)
      
      run  aquot toRun aqbq d aquot
    } else {
      toRun := RegExReplace(toRun, "\[(.+?)\]","") ; remove [...]
      d := modifyAttachmentsPath(attachment)
      
      run aquot toRun aqbq extractedExternalPath[1] aqbq d aquot
    }
    return
  }
}
;----------------------------------- exit -----------------------------------
exit(*){
  global
  
  OnMessage 0x03, moveEventSwitch, 0
  OnMessage 0x200, WM_MOUSEMOVED, 0
  
  if (pToken)
    Gdip_Shutdown(pToken)
  
  voiceIsSpeaker := 1
  voiceIsSpeed := 2 ; -10 .. +10
  sp := "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='de-DE'>"
  sp .= "Vielen Dank für die Verwendung von kontaktsupporter"
  sp .= "</speak>"
  speak(sp)
  
  sleep 500
  
  ExitApp
}
;----------------------------------------------------------------------------



