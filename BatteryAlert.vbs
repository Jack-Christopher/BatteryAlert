Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

Set BFCCResults = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")
For Each objItem in BFCCResults
	FullChargedCapacity = objItem.FullChargedCapacity
    WScript.Echo "FullChargedCapacity: " & FullChargedCapacity
Next

' The UpperLimit is the % of battery capacity that the system will 
' charge to before it warns the user that the battery is almost full
UpperLimit = FullChargedCapacity * 0.95

' The LowerLimit is the % of battery capacity that the system will
' have to warn the user that the battery is almost empty
LowerLimit = FullChargedCapacity * 0.10

' The SleepTime is the time in seconds that the script will wait
' before checking the battery status again
SleepTime = 300000 ' 5 minutes in milliseconds by default

while (1)
    Set BSResults = objWMIService.ExecQuery("Select * From BatteryStatus")
    For Each BSItem in BSResults
        ' the charge rate
        ChargeRate = BSItem.ChargeRate
        ' if it's charging
        IsCharging = BSItem.Charging
        ' the remaining capacity
        RemainingCapacity = BSItem.RemainingCapacity

        if (IsCharging AND ChargeRate > 0) then
            if (RemainingCapacity = FullChargedCapacity) then
                WScript.Echo "Battery is full"
                SleepTime = 60000 ' 60 seconds 
            elseif (RemainingCapacity >= UpperLimit) then
                WScript.Echo "Battery is almost full"
            end if
        elseif (RemainingCapacity <= LowerLimit) then
            WScript.Echo "Battery is almost empty"
            SleepTime = 60000 ' 60 seconds
        end if        
    next

    wscript.sleep SleepTime 

wend