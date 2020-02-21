param (
    [Parameter(Mandatory=$true)][string]$newFileName,
    [Parameter(Mandatory=$true)][string]$folderToCopyFrom,
    [Parameter(Mandatory=$true)][string[]]$weekNumbersToProcess
)

try
{
    If([System.IO.Path]::GetExtension($newFileName) -ne ".csv")
    {
        $newFileName += ".csv"
    }

    Add-Content -Path $newFileName -Value 'Type,Date,Day,Location,Contact,UserName,LastName,FirstName,Course,CRN,Major,GrantStatus,Name,WeekNumber'

    If($folderToCopyFrom.Substring($folderToCopyFrom.Length - 1) -ne "\")
    {
        $folderToCopyFrom += "\"
    }
    
    $files = Get-ChildItem "$folderToCopyFrom*.xlsm"

    ForEach ($file in $files)
    {
        #parse/split filename
        #foreach weekNumber in weekNumbers
        #if(weekNumber % 2 -eq 0)
        #  look at fileSplit[3] where length is the last 1-2 characters (based on weekNumber.length)

        $excel = New-Object -ComObject Excel.Application
        $excel.visible = $false
        $excel.displayalerts = $false

        $employeeFile = $excel.Workbooks.Open($file.FullName, $false)

        ForEach ($sheet in $employeeFile.Worksheets)
        {
            ForEach($weekNumberToProcess in $weekNumbersToProcess)
            {
                $lengthOfSubstring = 1

                if($weekNumberToProcess.Length -eq 2)
                {
                    $lengthOfSubstring = 2
                }
            
                $sheetName = $sheet.Name -split "_"
                $weekNumber = $sheetName[0].SubString($sheetName[0].Length-$lengthOfSubstring)
                $weekType = $sheetName[0].SubString(0,2)

                $sheetIsDifferentNumberOfDigits = ($sheetName[0].Substring($sheetName[0].Length-2, 1) -eq "1") -AND ($lengthOfSubstring -eq 1)

                if($weekType.ToString() -eq "CL" -AND $weekNumber.ToString() -eq $weekNumberToProcess.ToString() -AND !$sheetIsDifferentNumberOfDigits)
                {
                    $numberOfRows = $sheet.UsedRange.rows.count 

                    $tutorNameFile = Split-Path -Path $employeeFile.FullName -Leaf -Resolve
                    $tutorNameArray = $tutorNameFile -split "_"
                    $tutorName = "$($tutorNameArray[1]) $($tutorNameArray[2])"

                    for ($i=2; $i -le $numberOfRows; $i++)
                    {
                        $type = $sheet.Range("A$i").text        #colType
                        $date = $sheet.Range("B$i").text        #colDate
                        $day = $sheet.Range("C$i").text         #colDay
                        $location = $sheet.Range("D$i").text    #colLocation
                        $contact = $sheet.Range("E$i").text     #colContact
                        $userName = $sheet.Range("F$i").text    #colUserName
                        $lastName = $sheet.Range("G$i").text    #colLastName
                        $firstName = $sheet.Range("H$i").text   #colFirstName
                        $course = $sheet.Range("I$i").text      #colCourse
                        $CRN = $sheet.Range("J$i").text         #colCRN
                        $major = $sheet.Range("K$i").text       #colMajor
                        $grantStatus = $sheet.Range("L$i").text #colGrantStatus

                        if($contact -ne '.' -and
                        -not([string]::IsNullOrEmpty($type) `
                        -and [string]::IsNullOrEmpty($date) `
                        -and [string]::IsNullOrEmpty($day) `
                        -and [string]::IsNullOrEmpty($location) `
                        -and [string]::IsNullOrEmpty($contact) `
                        -and [string]::IsNullOrEmpty($userName) `
                        -and [string]::IsNullOrEmpty($lastName) `
                        -and [string]::IsNullOrEmpty($firstName) `
                        -and [string]::IsNullOrEmpty($course) `
                        -and [string]::IsNullOrEmpty($CRN) `
                        -and [string]::IsNullOrEmpty($major) `
                        -and [string]::IsNullOrEmpty($grantStatus)))
                        {
                            $newLine = "`"$type`",`"$date`",`"$day`",`"$location`",`"$contact`",`"$userName`",`"$lastName`",`"$firstName`",`"$course`",`"$CRN`",`"$major`",`"$grantStatus`",`"$tutorName`",`"$weekNumberToProcess`""
                            $newLine | Add-Content -path $newFileName
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        $employeeFile.Close($false)
        Remove-Variable -Name employeeFile

        $excel.Quit()
        Remove-Variable -Name excel
    }
}
catch
{
    Add-Content "$newFileName.txt" $file.FullName
    Add-Content "$newFileName.txt" $Error
    Write-Output $file.FullName
    Write-Output $Error
}
finally
{
    try {$excel.Quit()} catch{}

    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}
