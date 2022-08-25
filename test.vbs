set ws=createobject("wscript.shell")
on error resume next
a=ws.regread("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Rockwell Software\FactoryTalk\Platform\DCOMAuthLevel")
if err.number<>0 then
err.clear
ws.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Rockwell Software\FactoryTalk\Platform\DCOMAuthLevel",5,"REG_DWORD"
end if