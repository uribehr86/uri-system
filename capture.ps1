Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$Screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$Bitmap = New-Object System.Drawing.Bitmap($Screen.Width, $Screen.Height)
$Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
$Graphics.CopyFromScreen(0, 0, 0, 0, $Bitmap.Size)
$Bitmap.Save('C:\uri system scan\uri-system\screen.png')
