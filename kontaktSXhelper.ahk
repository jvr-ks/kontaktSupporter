; kontaktSXhelper.ahk
; Part of kontaktSupporter.ahk

;----------------------------- autoAppendTxtExt -----------------------------
autoAppendTxtExt(s){
  if (!RegExMatch(s, "\..+?$")){
    s := s ".txt"
  }
  return s
}
;----------------------------- extractExtension -----------------------------
extractExtension(f){
  local found, r
  
  r := ""
    
  SplitPath f ,,, &r
  
  return "." r
}
;------------------------------ extractFileName ------------------------------
extractFileName(f){
; TODO handle UNC, tests
  local found, r, extractedFileName
  
  r := ""
    
  SplitPath f , &r

  return r
}
;-------------------------------- extractPath --------------------------------
extractPath(f){
; TODO handle UNC, tests
  local found, r, extractedPath

  r := ""
    
  SplitPath f ,, &r
  
  return r
}
;-------------------------- getRunningAppsAsString --------------------------
getRunningAppsAsString(){
  DetectHiddenWindows false
  
  runningApps := ""
  runningAppsList := WinGetList(,, "Program Manager")
  for key, idNumber in runningAppsList {
    idTitle := WinGetTitle(idNumber)
    runningApps .= idTitle . " | "
  }
  return runningApps
}
;-------------------------- detectNewRunningAppLoop --------------------------
detectNewRunningAppLoop(theOpenLink, theCurrentRunningApps){
  global
  local theAppId
  
  theAppId := 0
  
  run theOpenLink
  
  appSearchLoop:
  loop 20 {
    theAppId := detectNewRunningApp(theCurrentRunningApps)
    if (theAppId != 0)
      break appSearchLoop
    sleep externalAppOpenDelay
  }
  
  if (!WinExist("ahk_id " theAppId)){
    theAppId := 0
  } else {
    showHintColored("Ok, gefunden: " WinGetTitle(theAppId) " (" theAppId ") ")
  }
  
  return theAppId
}
;----------------------------- detectNewRunningApp -----------------------------
detectNewRunningApp(lastRunning){
  local found, idTitle, foundAppId, runningAppsList, key, idNumber
  
  DetectHiddenWindows false
  
  found := 0
  idTitle := ""
  foundAppId := 0

  runningAppsList := WinGetList(,, "Program Manager")
  for key, idNumber in runningAppsList {
    idTitle := WinGetTitle(idNumber)
    if (idTitle != ""){
      if (!InStr(lastRunning, idTitle, 0)){
        foundAppId := idNumber
        found := 1
      }
    }
  }
  
  return foundAppId
}
;------------------------------- pathToRelativ -------------------------------
pathToRelativ(s, pathInUse){
  global 
  local relativDirectory
  
  relativDirectory := ""
  if (s != ""){
    ; remove baseDirectory from the path
    relativDirectory := StrReplace(s, pathInUse, "") ; remove path if file is inside default dir
    if (SubStr(relativDirectory, 1, 1) == "\")
      relativDirectory := SubStr(relativDirectory, 2) ; remove leading backslash
  }
  
  return relativDirectory
}
;--------------------------- modifyAttachmentsPath ---------------------------
modifyAttachmentsPath(attachmentLocal := ""){
  global
  local d
  
  attachmentLocal := RegExReplace(attachmentLocal,"\(.*\)","") ; remove comment
  
  d := attachmentsPathInUse
  if (attachmentLocal != ""){
    if (InStr(attachmentLocal, ":")){ ; absolute path?
      d := attachmentLocal ; ignore attachmentsPathInUse in this case
    } else {
      d := attachmentsPathInUse "\" attachmentLocal
    }
  }

  return d
}



;----------------------------------- todo -----------------------------------
todo(p1,p2,*){

  showHintColored("Programmteil `"" p1 "`" ist leider noch in Arbeit ...! (Menüpunkt: " p2 ")", 5000)

}
;---------------------------------- refresh ----------------------------------
refresh(*){

  Reload
}
;----------------------------- coordsScreenToApp -----------------------------
coordsScreenToApp(n){
  global dpiCorrect
  
  r := 0
  if (dpiCorrect > 0)
    r := round(n / dpiCorrect)

  return r
}
;----------------------------- coordsAppToScreen -----------------------------
coordsAppToScreen(n){
  global dpiCorrect

  r := round(n * dpiCorrect)

  return r
}
;------------------------------ showHintColored ------------------------------
showHintColored(s := "", n := 3000, fg := "cFFFFFF", bg := "a900ff", newfont := "", newfontsize := ""){
  global
  local t
  
  if (newfont == "")
    newfont := fontName
    
  if (newfontsize == "")
    newfontsize := fontsize
    
  if (hintColored != 0)
    hintColored.Destroy()
    
  hintColored := Gui("+0x80000000")
  hintColored.SetFont("c" fg " s" newfontsize, newfont)
  hintColored.BackColor := bg
  hintColored.add("Text", , s)
  hintColored.Opt("-Caption")
  hintColored.Opt("+ToolWindow")
  hintColored.Opt("+AlwaysOnTop")
  hintColored.Show("center")
  
  if (n > 0){
    sleep(n)
    hintColored.Destroy()
  }
}
;---------------------------- showHintColoredTop ----------------------------
showHintColoredTop(s := "", n := 3000, fg := "cFFFFFF", bg := "a900ff", newfont := "", newfontsize := ""){
  global
  local t
  
  if (hintColored != 0)
    hintColored.Destroy()
    
  if (newfont == "")
    newfont := fontName
    
  if (newfontsize == "")
    newfontsize := fontsize
  
  hintColored := Gui("+0x80000000")
  hintColored.SetFont("c" fg " s" newfontsize, newfont)
  hintColored.BackColor := bg
  hintColored.add("Text", , s)
  hintColored.Opt("-Caption")
  hintColored.Opt("+ToolWindow")
  hintColored.Opt("+AlwaysOnTop")
  hintColored.Show("y10 xcenter")
  
  if (n > 0){
    sleep(n)
    hintColored.Destroy()
  }
}
;----------------------------------- speak -----------------------------------
speak(text){
  global

  if (voiceIsEnabled){
    voice := ComObject("SAPI.SpVoice")
    voice.Voice := voice.GetVoices().Item(voiceIsSpeaker)
    voice.Rate := voiceIsSpeed
    voice.Speak(text)
  }
  
  return
}
;----------------------------------------------------------------------------


















