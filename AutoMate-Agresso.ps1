﻿<#
Hello and welcome to my script to #FuckAgresso and the end of spending hours reporting time in a usless stupid fucking system that never works and is a piece of crap.
If you are a regular normal person this is FREE OF USE.
If this is a big corparation/company please email johan.samuelsson42@gmail.com and I'd happily let you use MY code for a fair price. (100Eur/User/Month, can be discussed)

The script is ment to be used in conjunction with tampermonkey and a javascript to fill in the private notes when you close a ticket.
If you find this script useful and want to give me a gift in form of a Motorcycle, a new car, 1 million dollars or possibly some candy or an ice cream from the snack machine feel free to do so.
#>

<#NOTES
https://adamtheautomator.com/selenium-powershell/

https://www.jesusninoc.com/11/05/simulate-key-press-by-user-with-sendkeys-and-powershell/
https://stackoverflow.com/questions/17849522/how-to-perform-keystroke-inside-powershell
https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa202943(v=office.10)?redirectedfrom=MSDN
#>

#This makes the
#[System.Windows.Forms.SendKeys]::SendWait('shit to type')
#work.
#Bare in mind having this typing usernames and passwords are incredebly unsecure BUT FUCK AGRESSO!
Add-Type -AssemblyName System.Windows.Forms

# Your working directory
$workingPath = Get-Location

# Add the working directory to the environment path.
# This is required for the ChromeDriver to work.
if (($env:Path -split ';') -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}

#Import Selenium to PowerShell using the Import-Module cmdlet.
Import-Module "$($workingPath)\WebDriver.dll"

#Import the export from TimeReportPrepCSV.ps1
$CSV = Import-Csv .\export.csv

#Also having the username and password upload to github is fucking stupid, here's a config file; USE IT!
Foreach ($i in $(Get-Content Config.txt)) {
    Set-Variable -Name $i.split("=")[0] -Value $i.split("=", 2)[1]
}

$Headers = @{
    'Username' = $Username
    'Password' = $Password
}

#Below will purge some more fucking errors that are stupid and taking up space
#Don't aske me why log-level=3 does it, i don't have a fekking clue.
#.\chromedriver.exe --help
$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument('log-level=3')
<#
This can also be an array
$ChromeOptions.AddArgument(@(
        'log-level=3',
        'headless',
        'and so on'))
#>

#Function to download the latest ChromeDriver, it didn't need to be a fucntion but fuck you.
#This might be moved to a separate script?
function GetChromeDriver {
    #Check the latest STABLE version of ChromeDriver
    $ChromeDriverVersion = Invoke-WebRequest "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
    $ChromeDriverVersion = $ChromeDriverVersion.content
    #Get the script location
    $ScriptLocation = Get-Location

    #Remove old chromedriver file
    if (Test-Path -Path "$ScriptLocation\chromedriver.exe") {
        taskkill /IM "chromedriver.exe" /F
        Remove-Item "$ScriptLocation\chromedriver.exe"
    }

    #Download the latest stable ChromeDriver, extract it, move exe, and delete unused folders
    #Write it, cut it, paste it, save it, load it, check it, quick rewrite it
    #Charge it, point it, zoom it, press it, snap it, work it, quick erase it
    #Sorry for the Daft Punk - Technologic refference.
    Invoke-WebRequest "https://chromedriver.storage.googleapis.com/$ChromeDriverVersion/chromedriver_win32.zip" -OutFile "$ScriptLocation\driver.zip"
    Expand-Archive "$ScriptLocation\driver.zip"
    Move-Item -Path "$ScriptLocation\driver\*.exe" -Destination "$ScriptLocation"
    Remove-Item "$ScriptLocation\driver"
    Remove-Item "$ScriptLocation\driver.zip"
}

$CheckChromeDriverRunning = $false
while ($CheckChromeDriverRunning -eq $false) {
    try {
        #TRY to - Create a new ChromeDriver Object instance.
        $ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
        #If $ChromeDriver is running carry on
        if ($ChromeDriver) {
            $CheckChromeDriverRunning = $true
            Write-Host "Initiating launch of Chromedriver, prepare for Launch in 3...2...1..." -ForegroundColor Green
        }
    }
    catch {
        #If $ChromeDriver is not running dowload it via the GetChromeDriver function.
        if (!$ChromeDriver) {
            Write-Host "Chromedriver is not running, it's either out of date or non existant, downloading now..." -ForegroundColor Yellow
            GetChromeDriver
        }
    }
}

#Get primary monitor size
$PrimaryMonitorWidth = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
$PrimaryMonitorHeight = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height

#Set the fkn size of the window.
$ChromeDriver.Manage().Window.Size = "$PrimaryMonitorWidth,$PrimaryMonitorHeight"

#Move the fkn window.
$ChromeDriver.Manage().Window.Position = "0,0"

#Launch a browser and go to URL
$ChromeDriver.Navigate().GoToURL('https://agresso.advania.se/ubwprod/ContentContainer.aspx?type=topgen&menu_id=TS294&activityStepId=1-1&argument=&client=20&showMode=home')

Start-Sleep 1

#Type Username, press TAB, type password, press Enter. #Security101 :D
[System.Windows.Forms.SendKeys]::SendWait($Username)
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait($Password)
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')

#Wait for element fuction, call with
#WaitFor(' JS PATH ')
function WaitFor {
    param (
        $Element
    )
    $Wait = $false
    while ($Wait -eq $false) {
        Start-Sleep 1
        try {
            if ($ChromeDriver.FindElementByCssSelector($Element)) {
                $Wait = $true
                Write-Host 'Element found'
            } 
        }
        catch {
            Write-Host 'Element not found'
            Write-Host $Element
        }
    }
}

write-host "Please select correct week in agresso, press any key in this terminal window when done." -ForegroundColor Yellow;
cmd /c pause | out-null;

#Loop through the dates on Agresso so correct times can be registerd at correct dates.
$AgressoDates = @()
for ($i = 1; $i -lt 32; $i++) {
    if ($i -lt 10) {
        $j = "{0:00}" -f $i
    }
    else {
        $j = $i
    }
    try {
        if ($ChromeDriver.FindElementByXPath("//*[contains(@title, '$j/$CurrentMonth')]").Text -match ('\d\d\/\d\d')) {
            #For some fucking reason this is a fucking hashtable, what the fuck? I mean there is hash so that's the goo... DON'T DO DRUGS KIDS!
            #Or maybe it's not a hashtable, i can't fucking remember.
            $AgressoDate = $Matches
            $AgressoDates += $AgressoDate.Values
        }
    }
    catch {
        Out-Null
        #Write-Host "Error message fucking purged."
    }
}

$totalRows = 0;
foreach ($row in $CSV) {
    $totalRows++
}

$i = 0;
foreach ($row in $CSV) {
    $ticket = $row.Ticket
    #$row.Company
    $name = $row.Name
    $shortDescription = $row.ShortDescription
    $CloseNote = $row.CloseNote
    #$row.Avtal
    $code = $row.Code
    $time = $row.Time
    #Gotta do some date converting.
    $date = $row.Date
    $month = [regex]::match($date, '-(\d\d)-(\d\d)').Groups[1].Value
    $day = [regex]::match($date, '-(\d\d)-(\d\d)').Groups[2].Value
    $date = $day + '/' + $month

    #Naming variables fucking sucks ass fuck me
    #Shit $i is already a counter in this loop and... fuck it, $j it is.
    #Also it starts on one because for some reason that's what it does in the code... #FUCKAGRESSO
    $j = 1
    foreach ($DateInAgresso in $AgressoDates) {
        if ($date -eq $DateInAgresso) {
            Write-Host "Date match found."
            break
        }
        else {
            $j++
        }
    }

    $description = $ticket + '-' + $name + ', #Beskrivning: ' + $shortDescription + ' #Stägnings notis: ' + $CloseNote

    #Click "Lägg till" (Adds a row in agresso)
    #$ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_buttons__newButton').Click()
    #Above does not work, once "Lägg till" scrolls out of view it can no longer be clicked, however sending enter key to the element works.
    #This is done via ".SendKeys([OpenQA.Selenium.Keys]::Enter)", unlike "[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')" this seemes to work in the background.
    WaitFor('#b_s89_g89s90_buttons__newButton');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_buttons__newButton').SendKeys([OpenQA.Selenium.Keys]::Enter);
    
    #Write "Deploj" code.
    WaitFor('#b_s89_g89s90_row' + $i + '_1574_Editor');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_1574_Editor').SendKeys($code);
    
    #Write "Aktivitet" (Tek)
    WaitFor('#b_s89_g89s90_row' + $i + '_1576_Editor');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_1576_Editor').SendKeys('TEK');
    
    #Write "Beskrivning" (Description)
    WaitFor('#b_s89_g89s90_row' + $i + '_description_i');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_description_i').SendKeys($description);

    #Click Body because agresso is fucking retarded.
    $ChromeDriver.FindElementByCssSelector('body').click();

    #Write Time.
    WaitFor('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys([OpenQA.Selenium.Keys]::Enter);
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys([OpenQA.Selenium.Keys]::Control + "a");
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys($time);

    #Click Body because agresso is fucking retarded.
    $ChromeDriver.FindElementByCssSelector('body').click();


    $i++;
    Write-Host $i " of " $totalRows " rows done!";
}