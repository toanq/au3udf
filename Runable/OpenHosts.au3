#RequireAdmin
Opt("ExpandVarStrings", 1)

$sHostsPath = "@WindowsDir@\System32\drivers\etc\hosts"
$sEditor = "@WindowsDir@\System32\notepad.exe"

Run("$sEditor$ $sHostsPath$")