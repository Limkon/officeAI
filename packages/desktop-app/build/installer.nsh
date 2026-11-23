; OfficeAI 安装程序定制：安装完成后自动集成 Office / WPS（证书信任、Microsoft
; Office 加载项、WPS 加载项、唤醒器），用户无需手动点任何按钮、手动装证书。
; 集成逻辑在应用内（OfficeAI.exe --integrate 无头执行），此处仅负责在安装
; 完成后调用它。全部步骤幂等，升级重装时重复执行也无妨。
!macro customInstall
  DetailPrint "OfficeAI: 正在自动集成 Office / WPS（证书、加载项、唤醒器）..."
  ExecWait '"$INSTDIR\OfficeAI.exe" --integrate'
  DetailPrint "OfficeAI: 集成完成。打开 Word / Excel / PowerPoint / WPS 后，功能区会出现「AI 助手」按钮。"
!macroend
