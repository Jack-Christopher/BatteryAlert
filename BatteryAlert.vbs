Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

Set BFCCResults = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")
For Each objItem in BFCCResults
	FullChargedCapacity = objItem.FullChargedCapacity
    WScript.Echo "FullChargedCapacity: " & FullChargedCapacity
Next

UpperLimit = FullChargedCapacity * 0.95

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
            elseif (RemainingCapacity >= UpperLimit) then
                WScript.Echo "Battery is almost full"
            end if
        end if
        
    next

    wscript.sleep 60000 ' sleep for 60000 milliseconds (60 seconds)

wend