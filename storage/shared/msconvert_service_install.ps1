\\host.lan\Data\nssm install msconvert C:\Users\Docker\AppData\Local\Programs\Python\Python312\python.exe
#\\host.lan\Data\nssm set msconvert AppParameters '-m flask run --host 0.0.0.0'
\\host.lan\Data\nssm set msconvert AppParameters '-m uvicorn main:app --port 5000 --host 0.0.0.0'
\\host.lan\Data\nssm set msconvert AppDirectory \\host.lan\Data\msconvert\
\\host.lan\Data\nssm set msconvert AppStdout \\host.lan\Data\msconvert\msconvert.log
\\host.lan\Data\nssm set msconvert AppStderr \\host.lan\Data\msconvert\msconvert.log
\\host.lan\Data\nssm set msconvert AppStopMethodSkip 6
\\host.lan\Data\nssm set msconvert AppStopMethodConsole 1000
\\host.lan\Data\nssm set msconvert AppThrottle 5000
\\host.lan\Data\nssm set msconvert start SERVICE_AUTO_START
\\host.lan\Data\nssm set msconvert Type SERVICE_WIN32_OWN_PROCESS
\\host.lan\Data\nssm start msconvert
