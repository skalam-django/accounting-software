$BiosVer=(gwmi win32_bios).version
$CompName= (Get-WmiObject -Class Win32_ComputerSystem).name   

$BiosVer | out-file -filepath C:\Users\tawheed\AppData\Roaming\vbaTemp\Comp.dat -append
