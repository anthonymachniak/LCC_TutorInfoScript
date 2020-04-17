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

    Add-Content -Path $newFileName -Value 'Type,Date,Day,Location,Contact,UserName,LastName,FirstName,Course,CRN,Major,GrantStatus,Name,WeekNumber,Supervisor,EmployeeClass'

    If($folderToCopyFrom.Substring($folderToCopyFrom.Length - 1) -ne "\")
    {
        $folderToCopyFrom += "\"
    }
    
    $files = Get-ChildItem "$folderToCopyFrom*.xlsm"
    $regexPattern = '[^0-9]'

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $excel.displayalerts = $false

    ForEach ($file in $files)
    {
        $fileName = $file.Name
        $tutorNameArray = $fileName -split "_"
        $tutorName = "$($tutorNameArray[1]) $($tutorNameArray[2])"

        Write-Host "Checking $fileName for $tutorName"

        $week1 = $tutorNameArray[3]
        $week1 = $week1 -replace $regexPattern,''
        $week2 = $tutorNameArray[4]

        if ($weekNumbersToProcess -contains $week1 -or $weekNumbersToProcess -contains $week2)
        {
            $employeeFile = $excel.Workbooks.Open($file.FullName, $false)

            $employeeClass = ""
            $supervisor = ""
            ForEach ($sheet in $employeeFile.Worksheets)
            {
                $sheetName = $sheet.Name -split "_"
                $weekType = $sheetName[0].SubString(0,2)

                if($weekType.ToString() -eq "WL")
                {
                    $employeeClass = $sheet.Range("O3").text
                    $supervisor = $sheet.Range("O5").text
                    break;
                }
            }

            ForEach ($sheet in $employeeFile.Worksheets)
            {
                $sheetName = $sheet.Name -split "_"
                $weekType = $sheetName[0].SubString(0,2)

                if($weekType.ToString() -eq "CL")
                {
                    ForEach($weekNumberToProcess in $weekNumbersToProcess)
                    {
                        $weekNumber = $sheetName[0] -replace $regexPattern

                        if($weekNumber.ToString() -eq $weekNumberToProcess.ToString())
                        {
                            $numberOfRows = $sheet.UsedRange.rows.count 

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
                                    $newLine = "`"$type`",`"$date`",`"$day`",`"$location`",`"$contact`",`"$userName`",`"$lastName`",`"$firstName`",`"$course`",`"$CRN`",`"$major`",`"$grantStatus`",`"$tutorName`",`"$weekNumberToProcess`",`"$supervisor`",`"$employeeClass`""
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
            }

            $employeeFile.Close($false)
            Remove-Variable -Name employeeFile
        }
    }

    $excel.Quit()
    Remove-Variable -Name excel
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