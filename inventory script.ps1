PowerShell.exe -windowstyle hidden { 
<#
Set-ExecutionPolicy -ExecutionPolicy unrestricted
Set-ExecutionPolicy -ExecutionPolicy  ByPass
#>

# CSV File Path
$outfile = "\\domain\folder\Inventory.csv"

#First Get Macaddress
$mac=(Get-WmiObject Win32_NetworkAdapterConfiguration | where {$_.ipenabled -EQ $true}).Macaddress | select-object -first 1 -ErrorAction SilentlyContinue 
$macFound = 0

#Import CSV to check for MAc-Address 
$csv = Import-Csv $outfile

foreach ($row in $csv) {
    if ($row.Mac -eq $mac) { $macFound=1 }
    else {$macFound=0}
    }

    if ($macFound -eq 1) { Exit}
    else{
                #Get IPAddress
                $ip=(get-WmiObject Win32_NetworkAdapterConfiguration|Where {$_.Ipaddress.length -gt 1}).IPAddress | Select-object -index 0 -ErrorAction SilentlyContinue
    
                #Get CPU-INFO
                $cpu=Get-WmiObject -Class Win32_Processor |Select name -ErrorAction SilentlyContinue

                #Get Windwos-Specificaitons
                $windows=Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue

                #Get SerialNumber
                $serial=Get-WmiObject Win32_Bios |select SerialNumber -ErrorAction SilentlyContinue

                #Get OS-INFO
                $os=Get-ComputerInfo -Property "os*" | select OSName, OsArchitecture -ErrorAction SilentlyContinue

                #Grabs the Monitor objects from WMI
                    $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID" -ErrorAction SilentlyContinue
    
                    #Creates an empty array to hold the data
                    $Monitor_Array = @()
    
                    #Takes each monitor object found and runs the following code:
                    ForEach ($Monitor in $Monitors) {
      
                      #Grabs respective data and converts it from ASCII encoding and removes any trailing ASCII null values
                      If ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName) -ne $null) {
                        $Mon_Model = ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)","")
                        $Monitor_Array += $Mon_Model
                      } else {
                        $Mon_Model = "Unknown"
                      }}

                #Get Date
                $date=Get-Date -ErrorAction SilentlyContinue

                #Save To CSV
                $NewLine = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}" -f $ip,$mac,$windows.Domain,$windows.Manufacturer,$windows.Model,$windows.Name,$windows.PrimaryOwnerName,$windows.TotalPhysicalMemory,$cpu.name,$serial.SerialNumber,$os.OsName,$os.OsArchitecture,$Monitor_Array[0],$Monitor_Array[1],$date
                $NewLine | add-content -path $outfile
      }
}