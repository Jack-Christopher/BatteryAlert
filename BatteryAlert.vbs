Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

while (1)
    Set results = objWMIService.ExecQuery("Select * From BatteryStatus")
    For Each BSItem in results
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