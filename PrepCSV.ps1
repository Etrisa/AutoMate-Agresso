#Import csv file
$CustomerCodes = Import-Csv CustomerCodes.csv;
$import = Import-Csv task.csv -Delimiter ';';

foreach ($row in $import) {
    $Ticket = $row.Number;
    $Company = $row.Company;
    $PrivateNote = $row."Private notes (Advania visible only)";

    #Get name from private note
    $PrivateNoteName = [regex]::match($PrivateNote, 'Name:(.*)', 'IgnoreCase').Groups[1].Value;
    $PrivateNoteName = $PrivateNoteName.Trim();


    #Get short description from private note
    $PrivateNoteShortDescription = [regex]::match($PrivateNote, 'Short Description:(.*)', 'IgnoreCase').Groups[1].Value;
    $PrivateNoteShortDescription = $PrivateNoteShortDescription.Trim();

    #Get close note from prive note
    $PrivateNoteCloseNote = [regex]::match($PrivateNote, 'Close Notes:(.*)', 'IgnoreCase').Groups[1].Value;
    $PrivateNoteCloseNote = $PrivateNoteCloseNote.Trim();

    #Get "avtal" from private note
    $PrivateNoteInomAvtal = [regex]::match($PrivateNote, 'Avtal.*:(.*)', 'IgnoreCase').Groups[1].Value;
    $PrivateNoteInomAvtal = $PrivateNoteInomAvtal.Trim();

    foreach ($row in $CustomerCodes) {
        if ($Company -eq $row.Company -and $PrivateNoteInomAvtal -eq $row.Avtal) {
            $CompanyCode = $row.Code;
        }
    }

    $Date = [regex]::match($PrivateNote, '\d{4}-\d{2}-\d{2}', 'IgnoreCase').Groups[0].Value;
    $Date = $Date.Trim();

    #Get Time from private note
    $PrivateNoteTime = [regex]::match($PrivateNote, 'Time.*:(.*)', 'IgnoreCase').Groups[1].Value;
    $PrivateNoteTime = $PrivateNoteTime.Trim();

    #Write-Host "date:"$Date
    [PSCustomObject]@{
        Ticket           = $Ticket
        Company          = $Company
        Name             = $PrivateNoteName
        ShortDescription = $PrivateNoteShortDescription
        CloseNote        = $PrivateNoteCloseNote
        Avtal            = $PrivateNoteInomAvtal
        Code             = $CompanyCode
        Date             = $Date
        Time             = $PrivateNoteTime
    } | Export-Csv export.csv -Delimiter ',' -notype -Append -Encoding UTF8
}
#And let us look at the result before closing the window.
write-host "Done";
cmd /c pause | out-null;