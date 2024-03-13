; kontaktSXguiPreview.ahk
; Part of kontaktSupporter.ahk

;-------------------------------- previewGui --------------------------------
previewGui(){
  global

  paddingLeft := 0
  paddingRight := 10
  paddingTop := 120
  paddingBottom := 60
  sciRAWwidth := 800
  sciRAWheight := 600
  
  guiPreview := Gui("+OwnDialogs MaximizeBox +Resize + AlwaysOnTop", "Vorschau (keine Änderungen in der Vorschau möglich!)")
  
  guiPreviewMenu := MenuBar()
  guiPreviewMenu.Add("♒→" "Vorschau beenden", guiPreview_Close)
  guiPreviewMenu.Add("🗙" "kontaktSupporter beenden ", exit)
  
  guiPreview.MenuBar := guiPreviewMenu
  
  guiPreview.SetFont("S" . fontsizePreview, fontNamePreview)
  
  guiPreview.Add("Edit", "x" clientTopX " y" clientTopY " r0 h0 w600", "dummy1") ; Focus dummy and size dummy
  
  guiPreview.Add("Text", "w600 x" clientTopX, "eMail Parameter: ")
  previewParam := guiPreview.Add("Text", "w600 x" clientTopX)
  
  guiPreview.Add("Text", "w600 x" clientTopX, "eMail Parameter (expanded): ")
  previewParamExpanded := guiPreview.Add("Text", "w600 x" clientTopX)
  
  SB := guiPreview.Add("StatusBar")
  SB.SetParts(300, 300)
  
  if (!disableScintilla){
    sciStartX := paddingLeft
    sciStarty := paddingTop

    sciWidth := sciRAWwidth - paddingLeft - paddingRight
    sciHeight := sciRAWheight - paddingTop - paddingBottom
    
    previewText := guiPreview.AddScintilla("x" sciStartX " y" sciStarty " w" sciWidth " h" sciHeight "  DefaultOpt")
    setLighTheme(previewText)
    ;setKSLexer(previewText)
    
    previewText.callback := sci_Change
    
    previewText.Doc.ptr := previewText.Doc.Create(500000+100)

    previewText.Tab.Use := false
    previewText.Tab.Width := 2
    previewText.Margin.Width := 50
    previewText.Margin.Type := 0
    
    previewText.AutoSizeNumberMargin := true
  } else {
    previewText := guiPreview.Add("Edit", "w800 h600 x" clientTopX, "")
  }
  
  guiPreview.show("Hide x" guiPreviewPosX " y" guiPreviewPosY " w" guiPreviewClientWidth " h" guiPreviewClientHeight)
  
  ;WinSetAlwaysOnTop -1, "A"
  
  guiPreview.OnEvent("Size", guiPreview_Size , 1)
  guiPreview.OnEvent("Close", guiPreview_Close)
  
  guiPreviewHwnd := guiPreview.Hwnd 
  ; WinMinimize "A" ; trigger size-event
  ; WinRestore "ahk_id " guiPreviewHwnd
  
  previewIsVisible := 0

}
;------------------------------- guiPreview_Size -------------------------------
guiPreview_Size(theGui, theMinMax, clientWidth, clientHeight,*) {
  global 
  local previewTextWidth, previewTextHeight
  
  if (theMinMax = -1)
      return
  
  guiPreviewClientWidth := clientWidth
  guiPreviewClientHeight := clientHeight
  
  IniWrite clientWidth, configFile, "gui", "guiPreviewClientWidth"
  IniWrite clientHeight, configFile, "gui", "guiPreviewClientHeight"
  
  if (disableScintilla){  
    previewTextWidth := (guiPreviewClientWidth - marginLeft - marginRight)
    previewTextHeight := (guiPreviewClientHeight - previewTextMarginHeight)
  
    previewText.Move(, , previewTextWidth, previewTextHeight)
  } else {
    sciRAWwidth := guiPreviewClientWidth
    sciRAWheight := guiPreviewClientHeight

    previewTextWidth := sciRAWwidth - paddingLeft - paddingRight
    previewTextHeight := sciRAWheight - paddingTop - paddingBottom

    previewText.Move(, , previewTextWidth, previewTextHeight)   
  }
}
;------------------------------ guiPreviewMove ------------------------------
guiPreviewMove(){
  global
  local posX, posY
  
  if (IsSet(guiPreview)){
    guiPreview.GetPos(&posX, &posY)
  
    if (posX != 0 && posY != 0){  
      guiPreviewPosX := max(coordsAppToScreen(posX), minPosLeft)
      guiPreviewPosY := max(coordsAppToScreen(posY), minPosTop)

      IniWrite guiPreviewPosX, configFile, "gui", "guiPreviewPosX"
      IniWrite guiPreviewPosY, configFile, "gui", "guiPreviewPosY"
    }
  }
}
;----------------------------- guiPreview_Close -----------------------------
guiPreview_Close(*){
  global
  
  if (previewIsVisible){
    previewIsVisible := 0
    displayPreview := 0
    if (IsSet(guiPreview))
      guiPreview.Hide()
  }
}
;-------------------------------- sci_Change --------------------------------
sci_Change(*){
  global 
  local guiCtrlObj, CurrentCol, CurrentLine, oSaved

  guiCtrlObj := guiPreview.FocusedCtrl
  if (IsObject(guiCtrlObj)){
    CurrentCol := EditGetCurrentCol(guiCtrlObj)
    CurrentLine := EditGetCurrentLine(guiCtrlObj)
    ;oSaved := guiPreview.Submit(0)
    ;currentLineContent := guiPreview.Text
    ;tooltip currentLineContent
    SB.SetText("Line: " CurrentLine " Column: " CurrentCol , 2, 1)
  }
}
;------------------------------- setLighTheme -------------------------------
setLighTheme(ctl) {
    ; Light theme

  ctl.cust.Caret.LineBack := 0xF6F9FC	; active line (with caret)
  ctl.cust.Editor.Back := 0xFDFDFD

  ctl.cust.Editor.Fore := 0x000000
  ctl.cust.Editor.Font := "Consolas"
  ctl.cust.Editor.Size := 10

  ctl.Style.ClearAll()	; apply style 32

  ctl.cust.Margin.Back := 0xF0F0F0
  ctl.cust.Margin.Fore := 0x000000

  ctl.cust.Caret.Fore := 0x00FF00
  ctl.cust.Selection.Back := 0x398FFB
  ctl.cust.Selection.ForeColor := 0xFFFFFF

  ctl.cust.Brace.Fore := 0x5F6364	; basic brace color
  ctl.cust.BraceH.Fore := 0x00FF00	; brace color highlight
  ctl.cust.BraceHBad.Fore := 0xFF0000	; brace color bad light
  ctl.cust.Punct.Fore := 0xA57F5B
  ctl.cust.String1.Fore := 0x329C1B	; "" double quoted text
  ctl.cust.String2.Fore := 0x329C1B	; '' single quoted text

  ctl.cust.Comment1.Fore := 0x7D8B98	; keeping comment color same
  ctl.cust.Comment2.Fore := 0x7D8B98	; keeping comment color same
  ctl.cust.Number.Fore := 0xC72A31

  ctl.cust.kw1.Fore := 0x329C1B	; flow - set keyword list colors, kw1 - kw8
  ctl.cust.kw2.Fore := 0x1049BF	; funcion
  ctl.cust.kw2.Bold := true	; funcion
  ctl.cust.kw3.Fore := 0x2390B6	; method
  ctl.cust.kw3.Bold := true	; funcion
  ctl.cust.kw4.Fore := 0x3F8CD4	; prop
  ctl.cust.kw5.Fore := 0xC72A31	; vars

  ctl.cust.kw6.Fore := 0xEC9821	; directives
  ctl.cust.kw7.Fore := 0x2390B6	; var decl
}
;-------------------------------- rowPreview --------------------------------
rowPreview(*){
  global

  displayPreview := !displayPreview
  
  if (displayPreview)
    sendSelected(1)
  else
    guiPreview_Close()
    
}
;------------------------------ rowPreviewShow ------------------------------
rowPreviewShow(){
  global
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    if (!previewIsVisible){
      previewIsVisible := 1
      guiPreview.show()
    } else {
      msgbox "Fehler! Dieser Codeteil sollte nicht erreicht werden!"
    }
  }
}
;----------------------------- rowPreviewSetText -----------------------------
rowPreviewSetText(s){
  global
  
  if (!disableScintilla){
    previewText.Text := s
  } else {
    previewText.Value := s
  }
}
;--------------------------- generatePreviewParam ---------------------------
generatePreviewParam(){
  global
  
  param := format("{:03}", currentLineNumber) ": "
  param .= (nameRaw != "") ? nameRaw " / " : ""
  param .= (adrTextRaw != "") ? adrTextRaw " / " : ""
  param .= (ccTextRaw != "") ? ccTextRaw " / " : ""
  param .= (bccTextRaw != "") ? bccTextRaw " / " : ""
  param .= (subjectTextRaw != "") ? subjectTextRaw " / " : ""
  param .= (attachmentRaw != "") ? attachmentRaw " / " : ""
  
  previewParam.Value := param
  
  paramExpanded := format("{:03}", currentLineNumber) ": "
  paramExpanded .= (name != "") ? name " / " : ""
  paramExpanded .= (adrText != "") ? adrText " / " : ""
  paramExpanded .= (ccText != "") ? ccText " / " : ""
  paramExpanded .= (bccText != "") ? bccText " / " : "---"
  paramExpanded .= (subjectText != "") ? subjectText " / " : ""
  paramExpanded .= (attachment != "") ? attachment " / " : ""
  
  previewParamExpanded.Value := paramExpanded
}
;------------------------------- imagePreview -------------------------------
imagePreview(){
  global
  local ext, guiImagePreview, deltaX, deltaY, pBitmap, originalWidth, originalHeight, ratio, newWidth, newHeight
  local G, hbm, hdc, obm
  
  if (!imagePreviewIsVisible){
    if (attachment != ""){
      ext := extractExtension(extractFileName(attachment))
      if (ext = ".png" || ext = ".jpg" || ext = ".gif"){
        guiImagePreview := Gui("+toolwindow +AlwaysOnTop", attachment)
        guiImagePreview.OnEvent("Close", guiImagePreview_Close)
        
        If (!pToken) {
          guiImagePreview.Add("ActiveX", Format("w{1} h{2}", A_ScreenWidth - 30, A_ScreenHeight - 30), "mshtml:<img src='" attachment "' />")
        } else {
          ; GDI
          pBitmap := Gdip_CreateBitmapFromFile(attachment)
          If (!pBitmap){
            msgbox("Datei " attachment " konnte nicht geladen werden!", " Fehler aufgetreten!", "Icon!")
            return
          }
          previewImage := guiImagePreview.Add("Picture")
          ;previewImage.Value := Format("*w{1} *h{2} {3}", A_ScreenWidth, -1, attachment)
                  
          pBitmap := Gdip_CreateBitmapFromFile(attachment)
          newWidth := Gdip_GetImageWidth(pBitmap)
          newHeight := Gdip_GetImageHeight(pBitmap)

          sizeW := A_ScreenWidth - 30
          sizeH := A_ScreenHeight - 30
          if (newWidth > sizeW || newHeight > sizeH){
            resizeFactor := Min(sizeW/newWidth, sizeH/newHeight) 
                    
            PBitmapResized := Gdip_CreateBitmap(Round(newWidth * resizeFactor), Round(newHeight * resizeFactor))
            G := Gdip_GraphicsFromImage(pBitmapResized)
            Gdip_DrawImage(G, pBitmap, 0, 0, Round(newWidth * resizeFactor), Round(newHeight * resizeFactor), 0, 0, newWidth, newHeight)

            hCRBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmapResized)

            previewImage.Value := "HBITMAP:*" hCRBitmap

            Gdip_DeleteGraphics(G)
            Gdip_DisposeImage(PBitmapResized)
            Gdip_DisposeImage(pBitmap)
            DeleteObject(hCRBitmap)
          } else {
            hCRBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
            previewImage.Value := "HBITMAP:*" hCRBitmap
            Gdip_DisposeImage(pBitmap)
            DeleteObject(hCRBitmap)
          }
      }
      imagePreviewIsVisible := 1
      guiImagePreview.Show("autosize center")
      }
    }
  }
}
;--------------------------- guiImagePreviewClose ---------------------------
guiImagePreviewClose(*){
  global
  
  if (imagePreviewIsVisible){
    if (IsSet(guiImagePreview))
      guiImagePreview.Destroy()
  }
  imagePreviewIsVisible := 0
}
;--------------------------- guiImagePreview_Close ---------------------------
guiImagePreview_Close(*){
  global
  
  imagePreviewIsVisible := 0
}
;----------------------------------------------------------------------------
































