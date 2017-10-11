import sys
import subprocess as sp

ADLcmd = [
	r"C:\Installed Soft\FD\Tools\sdk\bin\adl.exe", 
	"application.xml", 
	"bin"
]

argLen = len(sys.argv)
if argLen > 1:
	ADLcmd.append("--")
	for x in range(1, argLen):
		ADLcmd.append(sys.argv[x])

startupinfo = sp.STARTUPINFO()
startupinfo.dwFlags |= sp.STARTF_USESHOWWINDOW

p = sp.Popen(ADLcmd, stdout=sp.PIPE, startupinfo=startupinfo)