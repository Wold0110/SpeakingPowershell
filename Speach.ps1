Add-Type -AssemblyName System.speech;
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;

#to set volume of output
Function Set-Speaker($Volume){$wshShell = new-object -com wscript.shell;1..50 | % {$wshShell.SendKeys([char]174)};1..$Volume | % {$wshShell.SendKeys([char]175)}}

$voicePacks =  $speak.GetInstalledVoices();
#foreach($voicePack in $voicePacks){
#    $vInfo = $voicePack.VoiceInfo
#    write-host $vInfo.Name;
#    write-host $vInfo.Age;
#    write-host $vInfo.Gender;
#    write-host $vInfo.Description;
#    write-host $vInfo.ID;
#}

$voiceIndex = Get-Random -Maximum $($voicePacks.Count)
write-host "Number of pack(s) installed:" $voicePacks.Count
$voiceSelected = $voicePacks[$voiceIndex].VoiceInfo.Name;
write-host "Voice used:"$voiceSelected
Set-Speaker 50 # X => 2X% | 50 -> 100
$speak.SelectVoice($voiceSelected);
$speak.Speak("Powershell speaking");
