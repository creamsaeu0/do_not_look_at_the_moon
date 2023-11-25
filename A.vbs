Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "C:\Users\이 문구를 지우고 여기에 사용자 이름을 입력해 주십시오.\Desktop\달을 보지 마십시오\B.bat", "/c lodctr.exe /r" , "", "runas",0
x=msgbox("당신 주변의 물체는 현재 <알 수 없는 무언가>로 인해 변화가 일어났습니다.",0+48,"<경고> 계속하려면 확인을 누르십시오.")
x=msgbox("당신과 가까이 있거나 특히 <당신의 손이 방금전까지 닿았던> 물체에 변화가 일어났을 확률이 매우 높습니다.",0+48,"<경고> 계속하려면 확인을 누르십시오.")
x=msgbox("당신 주변에 무언가 깜빡이지 않습니까?",0+32,"<경고> 깜빡.. 깜빡..")
