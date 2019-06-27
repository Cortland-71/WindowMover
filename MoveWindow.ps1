
$wshell = New-Object -ComObject wscript.shell;
$app = Start-Process "notepad"
Sleep 3
$wshell.AppActivate($app)
Sleep 1
$wshell.SendKeys('%' + " ")
Sleep .5
$wshell.SendKeys('{DOWN}')
Sleep .5
$wshell.SendKeys('~')
Sleep .5
$wshell.SendKeys('{RIGHT}')

$wshell.SendKeys('{RWin down}')