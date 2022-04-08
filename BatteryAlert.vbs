Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

Set BFCCResults = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")
For Each objItem in BFCCResults
	FullChargedCapacity = objItem.FullChargedCapacity
    WScript.Echo "FullChargedCapacity: " & FullChargedCapacity
Next

while (1)
    Set BSResults = objWMIService.ExecQuery("Select * From BatteryStatus")
    For Each BSItem in BSResults
        ' the charge rate
        WScript.Echo "ChargeRate:" & BSItem.ChargeRate
        ' if it's charging
        WScript.Echo "Charging:" & BSItem.Charging
        ' the discharge rate
        WScript.Echo "DischargeRate:" & BSItem.DischargeRate
        ' if it's discharging
        WScript.Echo "Discharging:" & BSItem.Discharging
        ' the remaining capacity
        WScript.Echo "RemainingCapacity:" & BSItem.RemainingCapacity
    next

    wscript.sleep 10000

wend