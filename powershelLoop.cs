LoopsPowershell.txt
loopsw in powershell schreiben: 

while Schleife

while ($variable = Read - Host - äö Prompt "Select a command") -ne "Q"){
	switch (§variable){
	   L {"File will be deleted"}
	   A {"displayed"}
	   R{"File will be write protected"}
	   Q{"End"}
	   default { "Invalid Entry"}
	 }
}


Ausführung ähnlich anderer Sprachen: Ausführung nur wenn die Konditionen stimmen.

do- while Schleife

do{
	loop body instructions
   }
   while/ until(condition)

do -while looped solange durch, bis die Kondition nicht mehr stimmt.
Do-Until looped durch bis sie stimmt.

(Demnach wäre die do-until-Schleife angebracht für das Projekt, wenn es um die Lizenzen verteilung geht.)
  

Break with GoTo

while(condition){
   loop body instructions
   if(condition) {break :DoLoop}
   loop body instructions
}
:DoLoop do{
   loop body instructions
}
until(condition)

Kann mit einem Label kombiniert werden. Wenn die if-Konditionen stimmen,
springt der Befehl zum Do weiter.

implicit for-each = Get-ADUser/* > (Get-ADUser -Filter *).Surname

reguläre for-each Schleife: 

$user = Get-ADUser -Filter *
foreach($u in $user) {
   $u.surname
}

das Alias für for each scheibt sich in Powershell als %

echo funktioniert wie bei PHP und gibt etwas aus.
if-Statement:
if($test) {echo "Value of test: " $test}

Regeln für conditional Statements:
Es muss nicht nach jeder Zeile ein Semicolon gesetzt werden, aber das Semicolon trennt innerhalb einer Zeile
mehrere Befehle. 

if(Test-Path *.gif){gci *.gif|foreach{$len += $_.length}; Write-Host $len " Bytes"}

if-else Statements sind anderen Sprachen gleich.

elseif Statements checken die anderen Konditionen, wenn das If-Statement false ausgibt, also nicht zutrifft. 

if (condition 1) {command}
elseif (condition 2) {command}
elseif (condition 3) {command}
else {command}

Switch Statement ( ohne CASE)

switch(Read-Host "Select a menu item"){
    1 {"File will be deleted"}
    2 {"File will be displayed"}
    3 {"File is write protected"}
    default {"Invalid entry"}
}

Powershell checkt alle Statements durch. Auch wenn eines bereits zuttrifft, werden die anderen geprüft.
Treffen dann noch welche zu, werden diese ebenfalls durchgeführt.
Um Strings miteinander zu vergleichen nutzt man CaseInsensitive und fügt in zu switch hinzu.
Weitere Operatoren zum vergleichen sind -wildcard und -regex.
Um zu verhindern dass die Befehle ausgeführt werden, muss am Ende eines jeden Befehls ein Break hinzugefügt
werden. Auf die Art endet der Befehl, bevor er weitere Befehle checkt.

switch -wildcard("PowerShell"){
    "Power*" {echo "'*' stands for 'shell'"}
    *ersh*" {echo "'*' replaces 'Pow' and 'ell'"}
    "PowerShe??" {echo "Pattern matches because ?? replaces two 'l' "}
}

