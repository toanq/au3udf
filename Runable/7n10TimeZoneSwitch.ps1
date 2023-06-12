$curr = (& $env:ComSpec /c tzutil /g)
if ($curr -eq 'SE Asia Standard Time')
{
    & $env:ComSpec /c tzutil /s "AUS Eastern Standard Time"
}
else
{
    & $env:ComSpec /c tzutil /s "SE Asia Standard Time"
}