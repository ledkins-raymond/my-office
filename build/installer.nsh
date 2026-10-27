; Adds a "My Excel" shortcut next to the "My Word" one electron-builder makes.
; Both start the same program; --app=excel opens a spreadsheet window.
!macro customInstall
  CreateShortCut "$SMPROGRAMS\My Excel.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=excel" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" 0
  CreateShortCut "$DESKTOP\My Excel.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=excel" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" 0
!macroend

!macro customUnInstall
  Delete "$SMPROGRAMS\My Excel.lnk"
  Delete "$DESKTOP\My Excel.lnk"
!macroend
