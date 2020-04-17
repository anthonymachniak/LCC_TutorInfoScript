param (
    [Parameter(Mandatory=$true)][string]$newFileName,
    [Parameter(Mandatory=$true)][string]$folderToCopyFrom
)

try
{
    If([System.IO.Path]::GetExtension($newFileName) -ne ".csv")
    {
        $newFileName += ".csv"
    }

    Add-Content -Path $newFileName -Value 'Date,Day,Time_Hrs,Students,TypeOfTimeSpent,TypeOfActivity,TypeOfGrant,EmployeeLastName,EmployeeFirstName'

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
        $employeeFirstName = $($tutorNameArray[1])
        $employeeLastName = $($tutorNameArray[2])

        $employeeFile = $excel.Workbooks.Open($file.FullName, $false)

        ForEach ($sheet in $employeeFile.Worksheets)
        {
            $sheetName = $sheet.Name -split "_"
            $weekType = $sheetName[0].SubString(0,2)

            if($weekType.ToString() -eq "WL")
            {
                $numberOfRows = $sheet.UsedRange.rows.count

                for ($i=19; $i -le $numberOfRows; $i++)
                {
                    $date = $sheet.Range("B$i").text        #colDate
                    $day = $sheet.Range("D$i").text         #colDay
                    $students = $sheet.Range("G$i").text    #colStudents
                    $timeSpent = $sheet.Range("H$i").text   #colTime
                    $activity = $sheet.Range("M$i").text    #colActivity
                    $grant = $sheet.Range("Q$i").text       #colGrant


                    $startTime = $sheet.Range("F$i").text
                    $endTime = $sheet.Range("E$i").text

                    $time_hrs = ""
                    if($startTime -eq "" -or $endTime -eq "")
                    {
                        $time_hrs = "N/A"
                    }
                    else
                    {
                        $timeDiff = New-TimeSpan $endTime $startTime
                        $time_hrs = $timeDiff.TotalHours
                    }

                    if(-not([string]::IsNullOrEmpty($date) `
                    -and [string]::IsNullOrEmpty($day) `
                    -and [string]::IsNullOrEmpty($students) `
                    -and [string]::IsNullOrEmpty($timeSpent) `
                    -and [string]::IsNullOrEmpty($activity) `
                    -and [string]::IsNullOrEmpty($grant)))
                    {
                        $newLine = "`"$date`",`"$day`",`"$time_hrs`",`"$students`",`"$timeSpent`",`"$activity`",`"$grant`",`"$employeeLastName`",`"$employeeFirstName`""
                        $newLine | Add-Content -path $newFileName
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }


        $employeeFile.Close($false)
        Remove-Variable -Name employeeFile
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