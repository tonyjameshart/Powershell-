If (!(Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue)) {
  If (!(Add-PSSnapin -PassThru VeeamPSSnapIn)) {
    Write-Error "Unable to load Veeam snapin" -ForegroundColor Red
    Exit
  }
}