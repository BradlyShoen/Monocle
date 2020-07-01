<#
	Name: Bradly Shoen
    Version: Monocle 1.2.0
	Description: Loops through a directory of images used for digital signage. Requires input from user to specify image directory location and duration of each display. Can be run through both GUI and command line.
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#This function resizes an image given via parameter based on the primary monitor size. The output will be an image that is either the same height, width, or both of the primary monitor.
Function ResizeImage() {
    param([String]$ImagePath)

    Add-Type -AssemblyName "System.Drawing"

    $img = [System.Drawing.Image]::FromFile($ImagePath)

    $CanvasWidth = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
    $CanvasHeight = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height

    #Encoder parameter for image quality
    $ImageEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($ImageEncoder, 90)

    #Get codec
    $Codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where {$_.MimeType -eq 'image/jpeg'}

    #Compute the final ratio to use
    $ratioX = $CanvasWidth / $img.Width;
    $ratioY = $CanvasHeight / $img.Height;

    $ratio = $ratioY
    if ($ratioX -le $ratioY) {
        $ratio = $ratioX
    }

    $newWidth = [int] ($img.Width * $ratio)
    $newHeight = [int] ($img.Height * $ratio)

    $bmpResized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graph = [System.Drawing.Graphics]::FromImage($bmpResized)
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    $graph.Clear([System.Drawing.Color]::White)
    $graph.DrawImage($img, 0, 0, $newWidth, $newHeight)

    $img.Dispose()

    return $bmpResized
}

#Initialize slides array. This variable will hold all of our resized images.
$global:slides = @()

#Initialize main Window form with no borders and maximized
$global:Window = New-Object system.Windows.Forms.Form
$global:Window.WindowState = 'Maximized'
$global:Window.FormBorderStyle = 'None'
$global:Window.TopMost = $true
$global:Window.BackColor = "Black"

#Initialize our timer that will control our slide transition
$timer = New-Object System.Windows.Forms.Timer
$timer.add_tick({
    
    #TIMER FUNCTIONALITY
    Write-Host $global:currentIndex
    if($global:slides.count -gt 1 -or $global:refreshedOnce -eq $false){
        #Remove our current slide from the Window and refresh canvas
        $global:Window.controls.remove($global:currentSlide)

        #Increase our current slide index and make sure we dont go out of bounds
        $global:currentIndex = $global:currentIndex + 1
        if($global:currentIndex -gt $global:slides.Count-1){
            $global:currentIndex = 0
        }

        #Grab image from our slides array based on the current slide index. Set our offsets so that the image will be centered vertically and horizontally.
        $global:img = $global:slides[$global:currentIndex]
        $global:horizontaloffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width - $global:img.Width)/2
        $global:verticaloffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height - $global:img.Height)/2

        #Update our current slide object to be the new image we just grabbed from slides.
        $global:currentSlide = new-object Windows.Forms.PictureBox
        $global:currentSlide.Height = $img.Size.Height
        $global:currentSlide.Width = $img.Size.Width
        $global:currentSlide.location = New-Object System.Drawing.Point($global:horizontaloffset,$global:verticaloffset)
        $global:currentSlide.Image = $img
        #Hide mouse cursor when hovering over the display and show when it is not
        $global:currentSlide.Add_MouseEnter({
            [System.Windows.Forms.Cursor]::Hide()
        })
        $global:currentSlide.Add_MouseLeave( {
            [System.Windows.Forms.Cursor]::Show()
        })

        #Add our current slide back to the Window and refresh canvas
        $global:Window.controls.add($global:currentSlide)
        $global:topSlide.SendToBack()
        $global:bottomSlide.SendToBack()
        $global:Window.Refresh()

        $global:refreshedOnce = $true
    }

})

#Initialize the Start Window
$StartWindow = New-Object system.Windows.Forms.Form
$StartWindow.ClientSize = '300,300'
$StartWindow.text = "Monocle By Bradly Shoen"
$StartWindow.TopMost = $true
$StartWindow.MaximizeBox = $false
$StartWindow.FormBorderStyle = 'Fixed3D'
$StartWindow.StartPosition = 'CenterScreen'

#This base64 string holds the bytes that make up the icon. Set our Start Window and Main Window form to our icon.
$iconBase64 = 'AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6AAAA2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFUAAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVQAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVAAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFUAAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVQAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVAAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFUAAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVQAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVAAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFUAAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVQAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVAAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAWAAAAawAAALoAAADdAAAA+AAAAOgAAADMAAAAiQAAADQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFUAAAD/AAAAAAAAAAAAAAAHAAAAfQAAAPcAAADuAAAApQAAAHIAAABbAAAAbAAA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'
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage = [System.Drawing.Image]::FromStream($stream, $true)
$StartWindow.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$global:Window.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

#Initialize our Start button and add functionality. Pressing button will start the Main Window form.
$Start = New-Object system.Windows.Forms.Button
$Start.text = "Start"
$Start.width = 60
$Start.height = 30
$Start.location = New-Object System.Drawing.Point(77,253)
$Start.Font = 'Microsoft Sans Serif,10'
$Start.Add_Click({
    
    #START BUTTON CODE

    #Grab every image at the specified slides directory (slides location textbox in GUI), resize the image, and add it to our slides array
    $files = Get-ChildItem ($SlidesLocation.Text+"\*") -Include *.jpg,*.jpeg,*.png,*.gif -ErrorAction SilentlyContinue
    for ($i=0; $i -lt $files.Count; $i++) {
        if($files[$i].extension -eq ".gif"){
            $newSlide = [Drawing.Image]::FromFile($files[$i].FullName)
        }else{
            $newSlide = ResizeImage($files[$i].FullName)
        }
        $global:slides = $global:slides + $newSlide
    }
    $global:currentIndex = 0
    

    #Create a file system watcher that will refresh our slides when a new file is added/deleted to the location folder so we don't have to restart the program
    $fsw = New-Object System.IO.FileSystemWatcher($SlidesLocation.Text, '*.*')
    $fsw.EnableRaisingEvents = $true
    $onCreated = Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action {
        $global:slides = @()
        $files = Get-ChildItem ($SlidesLocation.Text+"\*") -Include *.jpg,*.jpeg,*.png,*.gif -ErrorAction SilentlyContinue
        for ($i=0; $i -lt $files.Count; $i++) {
            if($files[$i].extension -eq ".gif"){
                $newSlide = [Drawing.Image]::FromFile($files[$i].FullName)
            }else{
                $newSlide = ResizeImage($files[$i].FullName)
            }
            $global:slides = $global:slides + $newSlide
        }
        $global:refreshedOnce = $false
        $global:currentIndex = 0
        
    }
    $onDeleted = Register-ObjectEvent $fsw Deleted -SourceIdentifier FileDeleted -Action {
        $global:slides = @()
        $files = Get-ChildItem ($SlidesLocation.Text+"\*") -Include *.jpg,*.jpeg,*.png,*.gif -ErrorAction SilentlyContinue
        for ($i=0; $i -lt $files.Count; $i++) {
            if($files[$i].extension -eq ".gif"){
                $newSlide = [Drawing.Image]::FromFile($files[$i].FullName)
            }else{
                $newSlide = ResizeImage($files[$i].FullName)
            }
            $global:slides = $global:slides + $newSlide
        }
        $global:refreshedOnce = $false
        $global:currentIndex = 0
    }

    #Setup top image and bottom image for vertical displays that would like stacked images
    if($args.Count -ge 3 -or $TopSlideLabel.Checked -eq $true){
        $topSlideImageRaw = Get-Item $TopSlide.Text -ErrorAction SilentlyContinue
        $topSlideImage = ResizeImage($topSlideImageRaw)
        $topSlideHorizontalOffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width - $topSlideImage.Width)/2
        $topSlideVerticalOffset = 0
        $global:topSlide = new-object Windows.Forms.PictureBox
        $global:topSlide.Image = $topSlideImage
        $global:topSlide.Width = $topSlideImage.size.Width-1
        $global:topSlide.Height = $topSlideImage.size.Height-1
        $global:topSlide.location = New-Object System.Drawing.Point($topSlideHorizontalOffset,$topSlideVerticalOffset)
        $global:topSlide.BorderStyle = "None"
        #Hide mouse cursor when hovering over the display and show when it is not
        $global:topSlide.Add_MouseEnter({
            [System.Windows.Forms.Cursor]::Hide()
        })
        $global:topSlide.Add_MouseLeave( {
            [System.Windows.Forms.Cursor]::Show()
        })
        $global:topSlide.SendToBack()
        $global:Window.controls.add($global:topSlide)
    }
    if($args.Count -ge 4 -or $BottomSlideLabel.Checked -eq $true){
        $bottomSlideImageRaw = Get-Item $BottomSlide.Text -ErrorAction SilentlyContinue
        $bottomSlideImage = ResizeImage($bottomSlideImageRaw)
        $bottomSlideHorizontalOffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width - $bottomSlideImage.Width)/2
        $bottomSlideVerticalOffset = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height-$bottomSlideImage.Height
        $global:bottomSlide = new-object Windows.Forms.PictureBox
        $global:bottomSlide.Image = $bottomSlideImage
        $global:bottomSlide.Width = $bottomSlideImage.size.Width-1
        $global:bottomSlide.Height = $bottomSlideImage.size.Height-1
        $global:bottomSlide.location = New-Object System.Drawing.Point($bottomSlideHorizontalOffset,$bottomSlideVerticalOffset)
        #Hide mouse cursor when hovering over the display and show when it is not
        $global:bottomSlide.Add_MouseEnter({
            [System.Windows.Forms.Cursor]::Hide()
        })
        $global:bottomSlide.Add_MouseLeave( {
            [System.Windows.Forms.Cursor]::Show()
        })
        $global:bottomSlide.SendToBack()
        $global:Window.controls.add($global:bottomSlide)
    }

    #Setup our initial image for our current slide and initialize our offsets to center the image vertically and horizontally.
    $global:img = $global:slides[$global:currentIndex]
    $global:horizontaloffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width - $img.Width)/2
    $global:verticaloffset = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height - $img.Height)/2

    #Initialize our current slide based on the image we grabbed and add it to the Window
    $global:currentSlide = new-object Windows.Forms.PictureBox
    $global:currentSlide.Height = $img.Size.Height
    $global:currentSlide.Width = $img.Size.Width
    $global:currentSlide.location = New-Object System.Drawing.Point($horizontaloffset,$verticaloffset)
    $global:currentSlide.Image = $img
    #Hide mouse cursor when hovering over the display and show when it is not
    $global:currentSlide.Add_MouseEnter({
        [System.Windows.Forms.Cursor]::Hide()
    })
    $global:currentSlide.Add_MouseLeave( {
        [System.Windows.Forms.Cursor]::Show()
    })
    $global:Window.controls.add($global:currentSlide)
    $global:currentSlide.BringToFront()

    #Set our timer interval (how fast our slides change) to be however long the user specified in the slides duration textbox in GUI in seconds.
    $timer.Interval = [int]$Duration.Text * 1000

    #Start our timer
    $timer.Start()

    #Display our main window.
    $global:Window.ShowDialog() | Out-Null

    #User has exited Main Window so we can close our Start Window.
    [void]$StartWindow.Close()
})

#Initialize our Cancel button and add functionality. Press button will close the program.
$Cancel = New-Object system.Windows.Forms.Button
$Cancel.text = "Exit"
$Cancel.width = 60
$Cancel.height = 30
$Cancel.location = New-Object System.Drawing.Point(184,253)
$Cancel.Font = 'Microsoft Sans Serif,10'
$Cancel.Add_Click({
    [void]$StartWindow.Close()
})

#Initialize our Slides Location Label ie text that says "Slide Location: "
$SlidesLabel = New-Object system.Windows.Forms.Label
$SlidesLabel.text = "Directory:"
$SlidesLabel.AutoSize = $true
$SlidesLabel.width = 25
$SlidesLabel.height = 10
$SlidesLabel.location = New-Object System.Drawing.Point(14,76)
$SlidesLabel.Font = 'Microsoft Sans Serif,10'

#Initialize Slides Location textbox ie where the user will specify where the folder is to grab the slides
$SlidesLocation = New-Object system.Windows.Forms.TextBox
$SlidesLocation.multiline = $false
$SlidesLocation.width = 166
$SlidesLocation.height = 20
$SlidesLocation.location = New-Object System.Drawing.Point(78,71)
$SlidesLocation.Font = 'Microsoft Sans Serif,13'
if($args.Count -ge 2){
    $SlidesLocation.Text = $args[0]
}else{
    $SlidesLocation.Text = "Example"
}

#Initialize our Slides Duration Label ie text that says "Slide Duration: "
$DurationLabel = New-Object system.Windows.Forms.Label
$DurationLabel.text = "Duration:"
$DurationLabel.AutoSize = $true
$DurationLabel.width = 25
$DurationLabel.height = 10
$DurationLabel.location = New-Object System.Drawing.Point(16,118)
$DurationLabel.Font = 'Microsoft Sans Serif,10'

#Initialize Slides Duration textbox ie where the user will specify how long the slides should take to transition
$Duration = New-Object system.Windows.Forms.TextBox
$Duration.multiline = $false
$Duration.width = 166
$Duration.height = 20
$Duration.location = New-Object System.Drawing.Point(77,114)
$Duration.Font = 'Microsoft Sans Serif,13'
$Duration.TextAlign = "Center"
$Duration.add_TextChanged({
    # Check if Text contains any non-Digits
    if($Duration.Text -match '\D'){
        # If so, remove them
        $Duration.Text = $tbox.Text -replace '\D'
        # If Text still has a value, move the cursor to the end of the number
        if($Duration.Text.Length -gt 0){
            $Duration.Focus()
            $Duration.SelectionStart = $Duration.Text.Length
        }
    }
})
if($args.Count -ge 2){
    $Duration.Text = $args[1]
}else{
    $Duration.Text = 5
}

#Initialize our Top Slide Checkbox
$TopSlideLabel = New-Object System.Windows.Forms.CheckBox
$TopSlideLabel.text = "Top:"
$TopSlideLabel.AutoSize = $true
$TopSlideLabel.width = 25
$TopSlideLabel.height = 10
$TopSlideLabel.location = New-Object System.Drawing.Point(16,159)
$TopSlideLabel.Font = 'Microsoft Sans Serif,10'
$TopSlideLabel.Add_CheckStateChanged({
    if($TopSlideLabel.Checked -eq $false){
        $TopSlide.Enabled = $false
    }else{
        $TopSlide.Enabled = $true
    }
})

#Initialize Top Slide textbox ie where the user will specify where the top slide image file is
$TopSlide = New-Object system.Windows.Forms.TextBox
$TopSlide.multiline = $false
$TopSlide.width = 136
$TopSlide.height = 20
$TopSlide.location = New-Object System.Drawing.Point(107,154)
$TopSlide.Font = 'Microsoft Sans Serif,13'
$TopSlide.TextAlign = "Left"
if($args.Count -ge 3){
    $TopSlide.Text = $args[2]
    $TopSlide.Enabled = $true
    $TopSlideLabel.Checked = $true
}else{
    $TopSlide.Enabled = $false
}

#Initialize our Bottom Slide Checkbox
$BottomSlideLabel = New-Object System.Windows.Forms.CheckBox
$BottomSlideLabel.text = "Bottom:"
$BottomSlideLabel.AutoSize = $true
$BottomSlideLabel.width = 25
$BottomSlideLabel.height = 10
$BottomSlideLabel.location = New-Object System.Drawing.Point(16,201)
$BottomSlideLabel.Font = 'Microsoft Sans Serif,10'
$BottomSlideLabel.Add_CheckStateChanged({
    if($BottomSlideLabel.Checked -eq $false){
        $BottomSlide.Enabled = $false
    }else{
        $BottomSlide.Enabled = $true
    }
})

#Initialize Bottom Slide textbox ie where the user will specify where the bottom slide image file is
$BottomSlide = New-Object system.Windows.Forms.TextBox
$BottomSlide.multiline = $false
$BottomSlide.width = 136
$BottomSlide.height = 20
$BottomSlide.location = New-Object System.Drawing.Point(107,196)
$BottomSlide.Font = 'Microsoft Sans Serif,13'
$BottomSlide.TextAlign = "Left"
if($args.Count -ge 4){
    $BottomSlide.Text = $args[3]
    $BottomSlide.Enabled = $true
    $BottomSlideLabel.Checked = $true
}else{
    $BottomSlide.Enabled = $false
}

#Initialize header
$headerImg = New-Object system.Windows.Forms.PictureBox
$headerImg.Width = 32
$headerImg.Height = 32
$headerImg.location = New-Object System.Drawing.Point(74,21)
$headerImg.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$headerImg.Image = $iconImage
$headerLabel = New-Object system.Windows.Forms.Label
$headerLabel.text = "Monocle"
$headerLabel.AutoSize = $true
$headerLabel.width = 25
$headerLabel.height = 10
$headerLabel.location = New-Object System.Drawing.Point(99,12)
$headerLabel.Font = 'Garamond,22,style=Bold'

#Check and see if the user presses escape during the display, if so close the program
$global:currentSlide_KeyDown = [System.Windows.Forms.KeyEventHandler]{

    if ($_.KeyCode -eq 'Escape')
    {
        [void]$global:Window.Close()
    }
}
$global:Window.Add_KeyDown($global:currentSlide_KeyDown)

#Hide mouse cursor when hovering over the display and show when it is not
$global:Window.Add_MouseEnter({
    [System.Windows.Forms.Cursor]::Hide()
})
$global:Window.Add_MouseLeave( {
    [System.Windows.Forms.Cursor]::Show()
})

#Add all labels, textboxes, and buttons to the Start Window form
$StartWindow.controls.AddRange(@($Start,$Cancel,$SlidesLabel,$SlidesLocation,$DurationLabel,$Duration,$headerImg,$headerLabel,$TopSlideLabel,$TopSlide,$BottomSlideLabel,$BottomSlide))

#Immediately press start button for console section
if($args.Count -ge 2){
    $StartWindow.add_Shown({
        $Start.PerformClick()
    })
}

#Display start window. Everything past this point is when the user closes the Start Window.
$StartWindow.ShowDialog() | Out-Null

#Cleanup our used variables
try{
    $global:Window.Dispose()
    $StartWindow.Dispose()
    $timer.Stop()
    $timer.Dispose()
    $global:currentSlide.Dispose()
    Unregister-Event -SourceIdentifier FileCreated -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier FileDeleted -ErrorAction SilentlyContinue
    $fsw.Dispose()
    $newSlide.Dispose()
}catch{}