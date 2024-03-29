echo "configure msoffice dcom"

while ((Get-WMIObject -Class Win32_DCOMApplicationSetting -Filter 'Description="Microsoft Word 97 - 2003 Document"') -eq $null)
{
    # powershell -NoProfile -ExecutionPolicy Bypass -Command "{}"
    # hack to reg ms office coms objects
    $wshell = New-Object -ComObject wscript.shell
    echo "clear"
    $wshell.SendKeys('{ESCAPE}')
    echo "start app"
    Start-Process dcomcnfg -PassThru
    Start-Sleep -Seconds 5
    Start-Sleep -Seconds 1
    echo "activate"
    $wshell.AppActivate("Component Services")
    Start-Sleep -Seconds 1
    echo "send keys"
    $wshell.SendKeys('{RIGHT}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{DOWN}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{RIGHT}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{DOWN}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{RIGHT}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{DOWN}')
    Start-Sleep -Seconds 1
    $wshell.SendKeys('{DOWN}')
    Start-Sleep -Seconds 5
    echo "done"

    taskkill /f /im mmc.exe
}

$apps = "Microsoft PowerPoint Slide", "Microsoft Excel Application", "Microsoft Graph Application", "Microsoft Visio previewer", "Microsoft Word 97 - 2003 Document"
Foreach ($app in $apps)
{
  echo ("processing {0}" -f $app)
  $dcom = Get-WMIObject -Class Win32_DCOMApplicationSetting -Filter ('Description="{0}"' -f $app)
  Set-ItemProperty -path ("Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\AppID\{0}" -f $dcom.AppID) -name "RunAs" -Value "Interactive User"
  echo ("{0}" -f $dcom)
}