﻿
# Author : JAGAT PRADHAN
#email : Jagat.Pradhan@live.com
# This script is intended to extract certain system details
# This will work for bulk system
#####CURRENT STATUS :::: WORK IN ROGRESS


 
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
    $serverInfoSheet.Cells.Item($row,$column)= 'WinDir'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'TimeZone'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'TotalMemory(GB)'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'Logical Processor'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= 'Free(C:)GB'
    $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Column++

    #Increment Row and reset Column back to first column
    $row++
    $Column = 1

    #Get the system details
    $list = Get-Content -Path 'D:\DO NOT DELETE_JAGAT\list.txt'


     foreach($server in $list)
      {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server
        $tz =  Get-CimInstance -ClassName Win32_TimeZone -ComputerName $server
        $tm = Get-CimInstance Win32_ComputerSystem -ComputerName $server | select name, @{Name ='TotalMemory';expression={[math]::Round($_.TotalPhysicalMemory / 1GB)}}
        $lp =  Get-CimInstance Win32_ComputerSystem -ComputerName $server | select NumberOfLogicalProcessors
        $vol =  Get-CimInstance Win32_Volume -ComputerName $server | select Name, @{Name = 'FreeSpace';expression={[math]::Round($_.FreeSpace / 1GB,2)}} | ?{$_.Name -eq 'C:\'}

        foreach($data in $os)
          {
       $serverInfoSheet.Cells.Item($row,$column)= $os.CSName
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $os.Caption
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $os.OSArchitecture
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $os.LastBootUpTime
       $column++
       $serverInfoSheet.Cells.Item($row,$column)= $os.WindowsDirectory
       $column++
      
       }

       foreach($value in $tz)
       {
       $serverInfoSheet.Cells.Item($row,$column)= $tz.Caption
       $Column++
       }

       foreach($msize in $tm)
       {
       $serverInfoSheet.Cells.Item($row,$column)= $tm.TotalMemory
       $Column++
        }
    foreach($psize in $lp)
       {

       $serverInfoSheet.Cells.Item($row,$column)= $lp.NumberOfLogicalProcessors
       $Column++
       }

       foreach($vsize in $vol)
       {
       $serverInfoSheet.Cells.Item($row,$column)= $vol.FreeSpace

          $row++
          $Column = 1
      
       }
    }
  }
 System_Details