<#
Hello and welcome to my script to #AutoMate Agresso and the end of spending hours reporting time stupid system that's a bit janky.
If you are a regular normal person this is FREE OF USE.
If you are corporation/company please email johan.samuelsson42 @ gmail.com and I'd happily let you use MY code for a fair price. (100Eur/User/Month, can be discussed)

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
#Bare in mind having this typing usernames and passwords are incredebly unsecure!
Add-Type -AssemblyName System.Windows.Forms

# Your working directory (Get script file location)
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

#Config file loop to import username and password for agresso!
Foreach ($i in $(Get-Content Config.txt)) {
    Set-Variable -Name $i.split("=")[0] -Value $i.split("=", 2)[1]
}

$Headers = @{
    'Username' = $Username
    'Password' = $Password
}

#Below will purge some errors that are taking up space
#Don't aske me why log-level=3 does it, I don't have a clue.
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

#Function to download the latest ChromeDriver.
function GetChromeDriver {
    #Check the latest STABLE version of ChromeDriver
    $ChromeDriverVersion = Invoke-WebRequest "https://chromedriver.storage.googleapis.com/LATEST_RELEASE" -UseBasicParsing
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
    Invoke-WebRequest "https://chromedriver.storage.googleapis.com/$ChromeDriverVersion/chromedriver_win32.zip" -OutFile "$ScriptLocation\driver.zip" -UseBasicParsing
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

#Get primary monitor size.
$PrimaryMonitorWidth = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
$PrimaryMonitorHeight = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height

#Set the size of the window.
$ChromeDriver.Manage().Window.Size = "$PrimaryMonitorWidth,$PrimaryMonitorHeight"

#Move the window.
$ChromeDriver.Manage().Window.Position = "0,0"

#Launch a browser and go to URL
$ChromeDriver.Navigate().GoToURL('https://agresso.advania.se/ubwprod/ContentContainer.aspx?type=topgen&menu_id=TS294&activityStepId=1-1&argument=&client=20&showMode=home')

#Take a nap.
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
                #Write-Host 'Element found'
            } 
        }
        catch {
            Write-Host 'Element not found, please contact Johan S'
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
            #For some reason this is a hashtable, wat?
            #Or maybe it's not a hashtable, i can't remember.
            #Anyways, convert this to something i can actually use.
            $AgressoDate = $Matches
            $AgressoDates += $AgressoDate.Values
        }
    }
    catch {
        Out-Null
        #Write-Host "Error message purged."
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

    #$i is already a counter in this loop $j it is.
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

    #Click Body because agresso needs to submit/commit the value and this is one way to do it.
    $ChromeDriver.FindElementByCssSelector('body').click();

    #Write Time.
    WaitFor('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i');
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys([OpenQA.Selenium.Keys]::Enter);
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys([OpenQA.Selenium.Keys]::Control + "a");
    $ChromeDriver.FindElementByCssSelector('#b_s89_g89s90_row' + $i + '_reg_value' + $j + '_i').SendKeys($time);

    #Click Body...
    $ChromeDriver.FindElementByCssSelector('body').click();

    $i++;
    Write-Host $i " of " $totalRows " rows done!";
}