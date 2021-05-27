
CD %USERPROFILE%\Documents\

curl -O https://raw.githubusercontent.com/TyGreenyy/CarOfferAHK/main/CarOfferAHK.ahk
curl -O https://www.autohotkey.com/download/ahk-install.exe

call %USERPROFILE%\Documents\ahk-install.exe /S

timeout 7

Call %USERPROFILE%\Documents\CarOfferAHK.ahk

