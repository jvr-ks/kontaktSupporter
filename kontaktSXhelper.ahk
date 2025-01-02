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
  
  run theOpenLink,,, &pid
  
  appName := ""
  try {
    appName := ProcessGetName(pid)
  } catch as e {
    showHintColored("eMail-App Name kann nicht bestimmt werden`n`nwhat: " e.what "`nfile: " e.file 
      . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra,, 16) 
  }
    
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
    showHintColored("Ok, neu gestartete eMail-App gefunden: " appName)
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
  
  d := attachmentsPathAbsolute
  if (attachmentLocal != ""){
    if (InStr(attachmentLocal, ":")){ ; absolute path?
      d := attachmentLocal ; ignore attachmentsPathAbsolute in this case
    } else {
      d := attachmentsPathAbsolute "\" attachmentLocal
    }
  }

  return d
}
;---------------------------- limitToViewableXPos ----------------------------
limitToViewableXPos(pos, objectWidth := 300){
  global

  ret := 0
  
  maxPosLeft := (A_ScreenWidth - objectWidth)
  minPosLeft := 0
  
  pos := min(pos, maxPosLeft)
  pos := max(pos, minPosLeft)
      
  return pos
}
;---------------------------- limitToViewableYPos ----------------------------
limitToViewableYPos(pos, objectHeight := 300){
  global

  ret := 0
  
  maxPosTop := (A_ScreenHeight - objectHeight)
  minPosTop := 0
  
  pos := min(pos, maxPosTop)
  pos := max(pos, minPosTop)
      
  return pos

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
showHintColored(s := "", n := 3000, fg := "cFFFFFF", bg := "a900ff", newfont := "", newfontsize := "", position := "center"){
  ; fontName and fontsize must be globally defined
  
  global
  local t
  
  if (newfont == "")
    newfont := fontName
    
  if (newfontsize == "")
    newfontsize := fontsize
    
  if (IsSet(hintColored))
    hintColored.Destroy()
    
  hintColored := Gui("+0x80000000 -Caption +ToolWindow +AlwaysOnTop")
  hintColored.SetFont("c" fg " s" newfontsize, newfont)
  hintColored.BackColor := bg
  hintColored.add("Text", , s)
  hintColored.Show(position)
  
  if (n > 0){
    ; delay the subsequent operations
    SetTimer destroyHintColored, -n
    sleep n
  } else {
    if (n != 0) ; don't destroy if n = 0
      ; don't delay the subsequent operations if n < 0
      SetTimer destroyHintColored, n
  }
}
;---------------------------- destroyHintColored ----------------------------
destroyHintColored(){
  global 
  
  hintColored.Destroy()
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
; =================================================================================
; From https://www.autohotkey.com/boards/viewtopic.php?style=1&f=83&t=114445
;
; Function: AutoXYWH
;   Move and resize control automatically when GUI resizes.
; Parameters:
;   DimSize - Can be one or more of x/y/w/h  optional followed by a fraction
;       add a '*' to DimSize to 'MoveDraw' the controls rather then just 'Move', this is recommended for Groupboxes
;       add a 't' to DimSize to tell AutoXYWH that the controls in cList are on/in a tab3 control
;   cList   - variadic list of GuiControl objects
;
; Examples:
;   AutoXYWH("xy", "Btn1", "Btn2")
;   AutoXYWH("w0.5 h 0.75", hEdit, "displayed text", "vLabel", "Button1")
;   AutoXYWH("*w0.5 h 0.75", hGroupbox1, "GrbChoices")
;   AutoXYWH("t x h0.5", "Btn1")
; ---------------------------------------------------------------------------------
; Version: 2023-02-25 / converted to v2 (Relayer)
;      2020-5-20 / small code improvements (toralf)
;      2018-1-31 / added a line to prevent warnings (pramach)
;      2018-1-13 / added t option for controls on Tab3 (Alguimist)
;      2015-5-29 / added 'reset' option (tmplinshi)
;      2014-7-03 / mod by toralf
;      2014-1-02 / initial version tmplinshi
; requires AHK version : v2+
; =================================================================================

AutoXYWH2(DimSize, cList*){   ;https://www.autohotkey.com/boards/viewtopic.php?t=1079

  static cInfo := map()

  if (DimSize = "reset")
  Return cInfo := map()

  for i, ctrl in cList {
    ctrlObj := ctrl
    ctrlObj.Gui.GetPos(&gx, &gy, &gw, &gh)
    if !cInfo.Has(ctrlObj) {
      ctrlObj.GetPos(&ix, &iy, &iw, &ih)
      MMD := InStr(DimSize, "*") ? "MoveDraw" : "Move"
      f := map( "x", 0
        , "y", 0
        , "w", 0
        , "h", 0 )

      for i, dim in (a := StrSplit(RegExReplace(DimSize, "i)[^xywh]"))) 
      if !RegExMatch(DimSize, "i)" . dim . "\s*\K[\d.-]+", &tmp)
        f[dim] := 1
      else
        f[dim] := tmp

      if (InStr(DimSize, "t")) {
      hWnd := ctrlObj.Hwnd
      hParentWnd := DllCall("GetParent", "Ptr", hWnd, "Ptr")
      RECT := buffer(16, 0)
      DllCall("GetWindowRect", "Ptr", hParentWnd, "Ptr", RECT)
      DllCall("MapWindowPoints", "Ptr", 0, "Ptr", DllCall("GetParent", "Ptr", hParentWnd, "Ptr"), "Ptr", RECT, "UInt", 1)
      ix := ix - NumGet(RECT, 0, "Int")
      iy := iy - NumGet(RECT, 4, "Int")
      }
      cInfo[ctrlObj] := {x:ix, fx:f["x"], y:iy, fy:f["y"], w:iw, fw:f["w"], h:ih, fh:f["h"], gw:gw, gh:gh, a:a, m:MMD}
    } else {
      dg := map( "x", 0
         , "y", 0
         , "w", 0
         , "h", 0 )
      dg["x"] := dg["w"] := gw - cInfo[ctrlObj].gw, dg["y"] := dg["h"] := gh - cInfo[ctrlObj].gh
      ctrlObj.Move(  dg["x"] * cInfo[ctrlObj].fx + cInfo[ctrlObj].x
         , dg["y"] * cInfo[ctrlObj].fy + cInfo[ctrlObj].y
         , dg["w"] * cInfo[ctrlObj].fw + cInfo[ctrlObj].w
         , dg["h"] * cInfo[ctrlObj].fh + cInfo[ctrlObj].h  )
      if (cInfo[ctrlObj].m = "MoveDraw")
      ctrlObj.Redraw()
    }
  }
}
;----------------------------------------------------------------------------


















