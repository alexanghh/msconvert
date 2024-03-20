echo "Install python..."
cp \\host.lan\Data\python-3.12.2-amd64.exe C:\Users\Docker\
~\python-3.12.2-amd64.exe /passive PrependPath=1
Start-Sleep -Seconds 5
rm -Force C:\Users\Docker\python-3.12.2-amd64.exe
echo "Install python completed"

echo "Install python packages..."
While (!(Test-Path C:\Users\Docker\AppData\Local\Programs\Python\Python312\Scripts\pip.exe -ErrorAction SilentlyContinue))
{
  # endless loop, when the file exist, it will continue
}
C:\Users\Docker\AppData\Local\Programs\Python\Python312\Scripts\pip.exe install (get-item \\host.lan\Data\python_packages\*.whl)
echo "Install python packages completed"

echo "Run fastapi to reg python on network..."
Start-Process 'C:\Users\Docker\AppData\Local\Programs\Python\Python312\python.exe' -WorkingDirectory '\\host.lan\Data\msconvert\' -ArgumentList '-m', 'uvicorn', 'main:app', '--host', '0.0.0.0'
Start-Sleep -Seconds 5
taskkill /f /im python.exe
echo "Run fastapi completed"

 echo "Setting python networking..."
$rules = Get-NetFirewallRule -All |? {$_.DisplayName -match "python.exe"}
$rules |% {
  Set-NetFirewallRule -DisplayName $_.DisplayName -Action Allow -Profile Any -Direction Inbound
}
Get-NetFirewallRule -All |? {$_.DisplayName -match "python.exe"}
echo "Setting python networking completed"