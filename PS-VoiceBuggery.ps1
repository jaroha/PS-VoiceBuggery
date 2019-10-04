## Load Speech Object
[Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null
## Load Sound Control
$volume = new-object -com wscript.shell
## Specify path to text file after '$phrasefile ='
## Text file should be formated by Seperating each phrase like so:
## "Phrase 1",
## "Phrase 2",
## "Phrase 3",
$phrasefile = "D:\text.txt"
## Just some error handling, if the file does not exist, spit out the random phrases below (they're stupid lyrics I found on a website)
if((Test-Path $phrasefile) -eq 0)
    { $phraselist = "I'm a little teapot short an stout, here is my handle, here is my spout","Me not working hard? Yea, right! Picture that with a Kodak And, better yet, go to Times Square Take a picture of me with a Kodak.”,“Only time will tell if we stand the test of time.” }
elseif ((Test-Path $phrasefile) -eq 1)
    { $phraselist = Get-Content $phrasefile }
#pull random phrase from either the file (if it exists) or the random phrases
$phrase = Get-Random $phraselist
## set the random speed
$ratenum = Get-Random -Maximum 4 -Minimum -5
$object = New-Object System.Speech.Synthesis.SpeechSynthesizer
## default speech voices installed in windows, additional ones can be downloaded from MS --
## leaving this in comment 
$voices = "Microsoft Hazel Desktop","Microsoft Zira Desktop","Microsoft David Desktop"
## Set Volume to max, no matter what
1..100 | % {$volume.SendKeys([char]175)}
##select the random voice
$voice = Get-Random $voices
$object.SelectVoice($voice)
$object.Rate = $ratenum
## say it
$object.Speak($phrase)