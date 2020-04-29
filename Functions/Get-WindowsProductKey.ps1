# create table to convert in base 24 
$map="BCDFGHJKMPQRTVWXY2346789" 
# Read registry Key 
$value = (get-itemproperty "HKLM:\\SOFTWARE\Microsoft\Windows NT\CurrentVersion").digitalproductid[0x34..0x42] 
# Convert in Hexa to show you the Raw Key 
$hexa = "" 
$value | foreach { 
  $hexa = $_.ToString("X2") + $hexa 
} 
"Raw Key Big Endian: $hexa" 
 
# find the Product Key 
$ProductKey = "" 
for ($i = 24; $i -ge 0; $i--) { 
  $r = 0 
  for ($j = 14; $j -ge 0; $j--) { 
    $r = ($r * 256) -bxor $value[$j] 
    $value[$j] = [math]::Floor([double]($r/24)) 
    $r = $r % 24 
  } 
  $ProductKey = $map[$r] + $ProductKey  
  if (($i % 5) -eq 0 -and $i -ne 0) { 
    $ProductKey = "-" + $ProductKey 
  } 
} 
"Product Key: $ProductKey"

function get-windowsproductkey([string]$computer)
{
$Reg = [WMIClass] ("\\" + $computer + "\root\default:StdRegProv")
$values = [byte[]]($reg.getbinaryvalue(2147483650,"SOFTWARE\Microsoft\Windows NT\CurrentVersion","DigitalProductId").uvalue)
$lookup = [char[]]("B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9")
$keyStartIndex = [int]52;
$keyEndIndex = [int]($keyStartIndex + 15);
$decodeLength = [int]29
$decodeStringLength = [int]15
$decodedChars = new-object char[] $decodeLength 
$hexPid = new-object System.Collections.ArrayList
for ($i = $keyStartIndex; $i -le $keyEndIndex; $i++){ [void]$hexPid.Add($values[$i]) }
for ( $i = $decodeLength - 1; $i -ge 0; $i--)
    {                
     if (($i + 1) % 6 -eq 0){$decodedChars[$i] = '-'}
     else
       {
        $digitMapIndex = [int]0
        for ($j = $decodeStringLength - 1; $j -ge 0; $j--)
        {
            $byteValue = [int](($digitMapIndex * [int]256) -bor [byte]$hexPid[$j]);
            $hexPid[$j] = [byte] ([math]::Floor($byteValue / 24));
            $digitMapIndex = $byteValue % 24;
            $decodedChars[$i] = $lookup[$digitMapIndex];
         }
        }
     }
$STR = ''     
$decodedChars | % { $str+=$_}
$STR
}
get-windowsproductkey .
