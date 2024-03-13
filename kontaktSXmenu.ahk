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
    menuItem4 := "Anhang (Dateien)"
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
;---------------------------- attachmentMenuBuild ----------------------------
attachmentMenuBuild(theMenu){
  global
  
  if (displayPreview){
    theMenu.Add("Datei auswählen", selectAttachmentFile)
    theMenu.Add("Ordner öffnen (mit dem Default-Filemanager)", openAttachmentFolder)
    theMenu.Add()
  }
  theMenu.Add("Externe Apps Liste `"" externalAppsFile "`" bearbeiten", editExternalApps)
  theMenu.Add()
  
  loop externalApps.length {
    theExternalName := externalApps[A_Index][1]
    if (theExternalName == "")
      theMenu.Add()
    else
      theMenu.Add(theExternalName, openExternal.Bind(externalApps[A_Index][1], externalApps[A_Index][2]))
  }
  
  return theMenu
}
;------------------------ attachmentMenuBuildRowEdit ------------------------
attachmentMenuBuildRowEdit(theMenu){
  global
  
  theMenu.Add("Datei auswählen", selectAttachmentFile)
  theMenu.Add("Ordner öffnen (mit dem Default-Filemanager)", openAttachmentFolder)
  theMenu.Add()

  theMenu.Add("Externe Apps Liste `"" externalAppsFile "`" bearbeiten", editExternalApps)
  theMenu.Add()

  theMenu := attachmentMenuBuild(theMenu)
  
  return theMenu
}
;----------------------------------------------------------------------------



































