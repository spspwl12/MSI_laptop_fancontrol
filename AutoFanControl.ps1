$wmi = Get-WmiObject MSI_ACPI -Namespace "root/wmi" -Filter "InstanceName='ACPI\\PNP0C14\\0_0'"
$prvjs = 0;
$tick = 9999;

foreach ($temp in $wmi) {
    $wmi = $temp;
}

$o = $wmi.Get_WMI();

function GetMethod
{
    param (
        [string]$Name,
        [byte[]]$Bytes
    )

    $p32 = $o.Clone();
    $p32val = $p32["Data"];
    $p32val.SetPropertyValue("Bytes", $Bytes);

    $rst = 0;

    if ($Name -eq "Thermal") {
        $rst = $wmi.Get_Thermal($p32val);
    } elseif ($Name -eq "Data") {
        $rst = $wmi.Get_Data($p32val);
    } elseif ($Name -eq "Fan") {
        $rst = $wmi.Get_Fan($p32val);
    } elseif ($Name -eq "AP") {
        $rst = $wmi.Get_AP($p32val);
    } elseif ($Name -eq "Temperature") {
        $rst = $wmi.Get_Temperature($p32val);
    }

    return $rst["Data"]["Bytes"];
}

function SetMethod
{
    param (
        [string]$Name,
        [byte[]]$Bytes
    )

    $p32 = $o.Clone();
    $p32val = $p32["Data"];
    $p32val.SetPropertyValue("Bytes", $Bytes);

    $rst = 0;

    if ($Name -eq "Data") {
        $rst = $wmi.Set_Data($p32val);
    } elseif ($Name -eq "Fan") {
        $rst = $wmi.Set_Fan($p32val);
    }

    return $rst["Data"]["Bytes"];
}

function GetTemperature {
    $nA = new-object byte[] 32;
    $Temp = GetMethod -Name "Temperature" -Bytes $nA;

    if ($Temp[0] -gt 0) {
        return [byte[]]($Temp[1], $Temp[2]);
    }

    return 0;
}

function GetFanRPM {
    $nA = new-object byte[] 32;
    $Temp = GetMethod -Name "Fan" -Bytes $nA;

    #$bytes = [byte[]]($rst[1], $rst[2], $rst[3])
    #[bitconverter]::ToInt16($bytes,0) 

    if ($Temp[0] -gt 0) {
        $f1 = [int](60000000.0 / ((( ($Temp[1] -shl 8) + $Temp[2]) * 2) * 62.5));
        $f2 = [int](60000000.0 / ((( ($Temp[3] -shl 8) + $Temp[4]) * 2) * 62.5));

        return [int[]]($f1, $f2);
    }

    return 0;
}

function AdvFan {
    param (
        [int]$Type,
        [byte]$Speed
    )

    $nA = new-object byte[] 32;
    $nA[0] = $Type;

    $rst = GetMethod -Name "Fan" -Bytes $nA;

    if ($rst[0] -gt 0) {

        $rst[0] = $Type;
        $rst[2] = $Speed; # FAN
        $rst[3] = $Speed;
        $rst[4] = $Speed;
        $rst[5] = $Speed;
        $rst[6] = $Speed;

        $rst = SetMethod -Name "Fan" -Bytes $rst;

        if ($rst[0] -gt 0) {

            $nA[0] = 1;
            $rst = GetMethod -Name "AP" -Bytes $nA;

            if ($rst[0] -gt 0) {

                $nA[0] = 212;
                $nA[1] = $rst[1];

                $nA[1] = $nA[1] -bor (1 -shl 6);
                $nA[1] = $nA[1] -band (-bnot (1 -shl 7));

                $rst = SetMethod -Name "Data" -Bytes $nA;

                if ($rst[0] -gt 0) {
                    return 1;
                }
            }
        }
    }

    return 0;
}

function CoolerBoost {
    param (
        [int]$IsEnabled
    )

    $nA = new-object byte[] 32;
    $nA[0] = 3;

    $rst = GetMethod -Name "Thermal" -Bytes $nA;

    if ($rst[0] -gt 0) {
        $nA[0] = 152;

        if ($IsEnabled -eq 0) {
            $nA[1] = $rst[1] -band (-bnot (1 -shl 7));
        } else {
            $nA[1] = $rst[1] -bor (1 -shl 7);
        }

        $rst = SetMethod -Name "Data" -Bytes $nA;

        if ($rst[0] -gt 0) {
            return 1;
        }
    }

    return 0;
}

function Set_FanSpeed {
    param (
        [int]$speed
    )

    $sp = [int]($speed / 9 * 150);

    AdvFan -Type 1 -Speed $sp > $nul;
    AdvFan -Type 2 -Speed $sp > $nul;
    CoolerBoost 0 > $nul;
}

function Justice {
    $js = 0;
    $pkg = GetTemperature;

    if ($pkg -eq 0) {
        return;
    }

    $ctemp = $pkg[0];
    $gtemp = $pkg[1];

    $temp = $ctemp;

    if ($gtemp -gt $ctemp) {
        $temp = $gtemp;
    }

    #Temperature Fan Speed Setting
    if ($temp -gt 68) { 
        $js = 10;
    } elseif ($temp -gt 65) {
        $js = 6;
    } elseif ($temp -gt 60) {
        $js = 5;
    } elseif ($temp -gt 55) {
        $js = 4;
    } elseif ($temp -gt 52) {
        $js = 3; 
    } elseif ($temp -gt 49) {
        $js = 2;
    } elseif ($temp -gt 46) {
        $js = 1;
    } elseif ($temp -gt 43) {
        $js = 1;
    } elseif ($temp -gt 40) {
        $js = 1;
    } elseif ($temp -gt 37) {
        $js = 1;
    } else {
        $js = 1;
    }

    if ($prvjs -eq $js) {
        return;
    }

    $prvjs = $js;

    if ($js -eq 10) {
        CoolerBoost 1;
        return;
    }

    Set_FanSpeed $js;
}

while ($true) {
    Justice; 
    Start-Sleep -Seconds 5;
}
