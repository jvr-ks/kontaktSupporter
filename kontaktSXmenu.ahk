; kontaktSXmenu.ahk
; Part of kontaktSupporter.ahk

;-------------------------------- setMenuTxt --------------------------------
setMenuTxt(){
  global
  
  if (menuTxtVisible){
    menuItem1 := "Bearbeiten"
    menuItem2 := "Vorschau"
    menuItem3 := "Tabelle"
    menuItem5 := "Konfiguration"
    menuItem4 := "Anhang (Global)"
    menuItem6 := "Hilfe"
    menuItem7 := "Refresh"
    menuItem8 := "[-...]"
    menuItem9 := "Exit"
  } else {
    menuItem1 := ""
    menuItem2 := ""
    menuItem3 := ""
    menuItem5 := ""
    menuItem4 := ""
    menuItem6 := ""
    menuItem7 := ""
    menuItem8 := "[+ ...]"
    menuItem9 := ""
  }
}
;------------------------------- toggleMenuTxt -------------------------------
toggleMenuTxt(*){
  global

  menuTxtVisible := !menuTxtVisible
  IniWrite menuTxtVisible, configFile, "gui", "menuTxtVisible"
  sleep 500
  Reload
}
;------------------------- attachmentMenuBuildGlobal -------------------------
attachmentMenuBuildGlobal(theMenu){
  global
  
  theMenu.Add("Globalen Ordner für Anhänge öffnen", openGlobalAttachmentFolder)
  theMenu.Add()
  
  theMenu.Add()
  
  loop externalApps.length {
    theExternalName := externalApps[A_Index][1]
    if (theExternalName == ""){
      theMenu.Add()
    } else {
      theMenu.Add("Ordner/Datei öffnen mit " theExternalName, openExternal.Bind(externalApps[A_Index][1], externalApps[A_Index][2]))
    }
  }
  
  theMenu.Add("Externe Apps Liste `"" externalAppsFile "`" bearbeiten", editExternalApps)
  
  return theMenu
}
;------------------------ openGlobalAttachmentFolder ------------------------
openGlobalAttachmentFolder(*){
  global
  local toRun
  
  toRun := attachmentsPathAbsolute
  
  if (InStr(attachmentsPath, ":")){ ; absolute path?
    toRun := extractPath(attachmentsPath)
  } else {
    if (InStr(attachmentsPath, "\")){
      found := RegExMatch(attachmentsPath,"(.*\\).*", &match) ; extract sub-directory
      if (found){
        toRun := attachmentsPathAbsolute "\" match.1
      } else {
        toRun := attachmentsPathAbsolute
        msgbox("Fehler beim bestimmen des Pfades für: `"" attachmentsPath "`" !", "Fehler aufgetreten!", "Icon!")
      }
    }
  }
  
  run toRun
}
;----------------------------------------------------------------------------



































