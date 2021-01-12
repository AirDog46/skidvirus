Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")

strDirectory = objShell.ExpandEnvironmentStrings("%temp%")

dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://images-ext-1.discordapp.net/external/tgIVDMaE7N7nUS49UhlCTdsSe4UdIi5HuwE6w__YyTw/%3Fresize%3D320%3A%2A/https/hips.hearstapps.com/hmg-prod.s3.amazonaws.com/images/james-charles-naked-butt-2020-1588945403.jpg", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile strDirectory + "\myImage.png", 2 '//overwrite
end with

objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", strDirectory + "\myImage.png"