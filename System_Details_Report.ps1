function System_Details()
{
        
    #Create excel COM object
    $excel = New-Object -ComObject excel.application

    #Make Visible
    $excel.Visible = $True

    #Add a workbook
    $workbook = $excel.Workbooks.Add()
        
    #Connect to first worksheet to rename and make active
    $serverInfoSheet = $workbook.Worksheets.Item(1)
    $serverInfoSheet.Name = 'System_Details'
    $serverInfoSheet.Activate() | Out-Null

    #Create a Title for the first worksheet and adjust the font
    $row = 1
    $Column = 1
    $serverInfoSheet.Cells.Item($row,$column)= 'System Details'

    # Making the tite bit larger and coloutful
    $serverInfoSheet.Cells.Item($row,$column).Font.Size = 18
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $serverInfoSheet.Cells.Item($row,$column).Font.Name = "Cambria"
    $serverInfoSheet.Cells.Item($row,$column).Font.ThemeFont = 1
    $serverInfoSheet.Cells.Item($row,$column).Font.ThemeColor = 4
    $serverInfoSheet.Cells.Item($row,$column).Font.ColorIndex = 55
    $serverInfoSheet.Cells.Item($row,$column).Font.Color = 8210719

    #Bit more clean up
    $range = $serverInfoSheet.Range("a1","g2")
    $range.Merge() | Out-Null
    $range.VerticalAlignment = -4160

    #Increment row for next set of data
    $row++;$row++
    
    #Save the initial row so it can be used later to create a border
    $initalRow = $row

    #Create a header for System Detais Report; set each cell to Bold and add a background color
    $serverInfoSheet.Cells.Item($row,$column)= 'Hostname'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'OS'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'OS-Arch'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'LastBootTime'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'LocalDate-Time'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'WinDir'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'TimeZone'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++

    #Increment Row and reset Column back to first column
    $row++
    $Column = 1

    #Get the system details
    $tz = Get-CimInstance -ClassName Win32_TimeZone
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    
    # Write the fetched data into rows & column
     
     $os | ForEach-Object {
       
       $serverInfoSheet.Cells.Item($row,$column)= $_.csname
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $_.Caption
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $_.OSArchitecture
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $_.LastBootUpTime
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $_.LocalDateTime
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $_.WindowsDirectory
       $column++
          
    }


    # timezone Exporting to Excel
      $tz | ForEach-Object {       
      $serverInfoSheet.Cells.Item($row,$column) = $_.Caption
      $column++
      }


           
 }

 System_Details