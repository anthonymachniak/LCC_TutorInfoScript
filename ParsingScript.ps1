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

    Write-Output $newFileName

    Add-Content -Path $newFileName -Value 'Type,Date,Day,Location,Contact,UserName,LastName,FirstName,Course,CRN,Major,GrantStatus,Name,WeekNumber'


    If($folderToCopyFrom.Substring($folderToCopyFrom.Length - 1) -ne "\")
    {
        $folderToCopyFrom += "\"
    }
    
    $files = Get-ChildItem "$folderToCopyFrom*.xlsm"

    Write-Output $folderToCopyFrom
    Write-Output $files

    ForEach ($file in $files)
    {
        $excel = New-Object -ComObject Excel.Application
        $excel.visible = $false
        $excel.displayalerts = $false

        Write-Output $file.FullName
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

                Write-Output $sheetIsDifferentNumberOfDigits
                Write-Output $weekNumber
                Write-Output $weekType

                if($weekType.ToString() -eq "CL" -AND $weekNumber.ToString() -eq $weekNumberToProcess.ToString() -AND !$sheetIsDifferentNumberOfDigits)
                {
                    $numberOfRows = $sheet.UsedRange.rows.count 

                    $colType = "A"
                    $colDate = "B"
                    $colDay = "C"
                    $colLocation = "D"
                    $colContact = "E"
                    $colUserName = "F"
                    $colLastName = "G"
                    $colFirstName = "H"
                    $colCourse = "I"
                    $colCRN = "J"
                    $colMajor = "K"
                    $colGrantStatus = "L"

                    $tutorNameFile = Split-Path -Path $employeeFile.FullName -Leaf -Resolve
                    Write-Output "tutorNameFile:" $tutorNameFile

                    $tutorNameArray = $tutorNameFile -split "_"
                    Write-Output "tutorNameArray:" $tutorNameArray

                    $tutorName = "$($tutorNameArray[1]) $($tutorNameArray[2])"
                    Write-Output "tutorName:" $tutorName

                    for ($i=2; $i -le $numberOfRows; $i++)
                    {
                        $type = $sheet.Range("$colType$i").text
                        $date = $sheet.Range("$colDate$i").text
                        $day = $sheet.Range("$colDay$i").text
                        $location = $sheet.Range("$colLocation$i").text
                        $contact = $sheet.Range("$colContact$i").text
                        $userName = $sheet.Range("$colUserName$i").text
                        $lastName = $sheet.Range("$colLastName$i").text
                        $firstName = $sheet.Range("$colFirstName$i").text
                        $course = $sheet.Range("$colCourse$i").text
                        $CRN = $sheet.Range("$colCRN$i").text
                        $major = $sheet.Range("$colMajor$i").text
                        $grantStatus = $sheet.Range("$colGrantStatus$i").text


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
                            $newLine = "`"{0}`",`"{1}`",`"{2}`",`"{3}`",`"{4}`",`"{5}`",`"{6}`",`"{7}`",`"{8}`",`"{9}`",`"{10}`",`"{11}`",`"{12}`",`"{13}`"" -f $type, $date, $day, $location, $contact, $userName, $lastName, $firstName, $course, $CRN, $major, $grantStatus, $tutorName, $weekNumberToProcess
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
