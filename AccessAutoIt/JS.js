
// Start AutoIt server script first

// Let's go...
try
{
    oAutoIt = GetObject("AutoIt.Application");
    oAutoIt.Call("Beep", 1000, 500)
	oAutoIt.Call("MsgBox", 64 + 262144, "Hi there!", "Is this something or what?")
	if (oAutoIt.Call("MsgBox", 4 + 48 + 262144, "?", "Kill server?") == 6)
	{
	    oAutoIt.Quit;
	}
}
catch(err)
{
    WScript.echo("!Ermmm... is server script runing?");
}



