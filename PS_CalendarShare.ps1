Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create form
$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Point(400,400)
$Form.Text = "Calendar Sharing Module"

# Create box to hold each page data. Starting with login page
$GroupboxLogin = New-Object System.Windows.Forms.GroupBox
$GroupboxLogin.Height = 450
$GroupboxLogin.width = 400
$GroupboxLogin.Location = New-Object System.Drawing.Point(0,0)

$ButtonInstall = New-Object System.Windows.Forms.Button
$ButtonInstall.Text = "Install module"
$ButtonInstall.Width = 60
$ButtonInstall.Height = 50
$ButtonInstall.Location = New-Object System.Drawing.Point(150,50)
# Installs the required module
$ButtonInstall.Add_Click({
    SystemModuleInstall
    [System.Windows.MessageBox]::Show('Module Installed')
})

$ButtonLogin = New-Object System.Windows.Forms.Button
$ButtonLogin.Text = "Login to system"
$ButtonLogin.Width = 60
$ButtonLogin.Height = 50
$ButtonLogin.Location = New-Object System.Drawing.Point(150,150)
# Generate click functionality, calls two functions called later in functions section
$ButtonLogin.Add_Click({
    SystemLogin
    SystemLoadMenu
})

# Closes system

$ButtonQuit = New-Object System.Windows.Forms.Button
$ButtonQuit.Text = "Close system"
$ButtonQuit.Width = 60
$ButtonQuit.Height = 50
$ButtonQuit.Location = New-Object System.Drawing.Point(150,250)
$ButtonQuit.Add_Click({
    SystemQuit
})


# Create main menu page

$GroupboxMenu = New-Object System.Windows.Forms.GroupBox
$GroupboxMenu.Height = 450
$GroupboxMenu.width = 400
$GroupboxMenu.Visible = $false
$GroupboxMenu.Location = New-Object System.Drawing.Point(0,0)

$ButtonGrant = New-Object System.Windows.Forms.Button 
$ButtonGrant.Text = "Grant calendar permissions"
$ButtonGrant.width = 75
$ButtonGrant.height = 50
$ButtonGrant.Location = New-Object System.Drawing.Point(150,50)
$ButtonGrant.Add_click({
    SystemShareCalendarMenu
    SystemHideMenu
})

$ButtonGet = New-Object System.Windows.Forms.Button 
$ButtonGet.Text = "Get calendar permissions"
$ButtonGet.width = 75
$ButtonGet.height = 50
$ButtonGet.Location = New-Object System.Drawing.Point(150,150)
$ButtonGet.Add_Click({
    SystemGetCalendarMenu
    SystemHideMenu
})

$ButtonRevoke = New-Object System.Windows.Forms.Button 
$ButtonRevoke.Text = "Revoke Calendar Permissions"
$ButtonRevoke.width = 75
$ButtonRevoke.height = 50
$ButtonRevoke.Location = New-Object System.Drawing.Point(150,250)
$ButtonRevoke.Add_click({
    SystemRevokeCalendarMenu
    SystemHideMenu
})

# Share Menu 

$GroupboxShare = New-Object System.Windows.Forms.GroupBox
$GroupboxShare.Height = 450
$GroupboxShare.width = 400
$GroupboxShare.Visible = $false
$GroupboxShare.Location = New-Object System.Drawing.Point(0,0)

$labelShareFrom = New-Object System.Windows.Forms.Label
$labelShareFrom.Text = "From user (User Calendar)"
$labelShareFrom.width = 200
$labelShareFrom.height = 20
$labelShareFrom.Location = New-Object System.Drawing.Point(130,100)

$InputShareFrom = New-Object System.Windows.Forms.TextBox
$InputShareFrom.height = 10
$InputShareFrom.width = 100
$InputShareFrom.Location = New-Object System.Drawing.Point(150,125)

$labelShareTo = New-Object System.Windows.Forms.Label 
$labelShareTo.Text = "To user (User gaining provider)"
$labelShareTo.width = 200
$labelShareTo.height = 20
$labelShareTo.Location = New-Object System.Drawing.Point(130,155)

$InputShareTo = New-Object System.Windows.Forms.TextBox
$InputShareTo.height = 10
$InputShareTo.width = 100
$InputShareTo.Location = New-Object System.Drawing.Point(150,175)

$ButtonShareCalendar = New-Object System.Windows.Forms.Button 
$ButtonShareCalendar.text = "Share Calendar"
$ButtonShareCalendar.width = 75
$ButtonShareCalendar.height = 50 
$ButtonShareCalendar.Location = New-Object System.Drawing.Point(150,200)
$ButtonShareCalendar.Add_Click({
    ShareCalendar
})

$ButtonReturnShare = New-Object System.Windows.Forms.Button 
$ButtonReturnShare.text = "Return to menu"
$ButtonReturnShare.width = 75
$ButtonReturnShare.height = 50
$ButtonReturnShare.Location = New-Object System.Drawing.Point(100,300)
$ButtonReturnShare.Add_Click({
    SystemReturn
})

$ButtonQuitShare = New-Object System.Windows.Forms.Button
$ButtonQuitShare.Text = "Close system"
$ButtonQuitShare.Width = 60
$ButtonQuitShare.Height = 50
$ButtonQuitShare.Location = New-Object System.Drawing.Point(200,300)
$ButtonQuitShare.Add_Click({
    SystemQuit
})

# Get menu

$GroupboxGet = New-Object System.Windows.Forms.GroupBox
$GroupboxGet.Height = 450
$GroupboxGet.width = 400
$GroupboxGet.Visible = $false
$GroupboxGet.Location = New-Object System.Drawing.Point(0,0)

$InputGetPermissions = New-Object System.Windows.Forms.TextBox
$InputGetPermissions.height = 10
$InputGetPermissions.width = 100
$InputGetPermissions.Location = New-Object System.Drawing.Point(150,175)

$ButtonGetPermissions = New-Object System.Windows.Forms.Button
$ButtonGetPermissions.text = "Get calendar permissions"
$ButtonGetPermissions.width = 100
$ButtonGetPermissions.height = 50
$ButtonGetPermissions.location = New-Object System.Drawing.Point(150,200)
$ButtonGetPermissions.Add_click({
    GetCalendarPermissions
})

$ButtonReturnGet = New-Object System.Windows.Forms.Button 
$ButtonReturnGet.text = "Return to menu"
$ButtonReturnGet.width = 75
$ButtonReturnGet.height = 50
$ButtonReturnGet.Location = New-Object System.Drawing.Point(100,300)
$ButtonReturnGet.Add_Click({
    SystemReturn
})

$ButtonQuitGet = New-Object System.Windows.Forms.Button
$ButtonQuitGet.Text = "Close system"
$ButtonQuitGet.Width = 60
$ButtonQuitGet.Height = 50
$ButtonQuitGet.Location = New-Object System.Drawing.Point(200,300)
$ButtonQuitGet.Add_Click({
    SystemQuit
})


# Revoke menu

$GroupboxRevoke = New-Object System.Windows.Forms.GroupBox
$GroupboxRevoke.Height = 450
$GroupboxRevoke.width = 400
$GroupboxRevoke.Visible = $false
$GroupboxRevoke.Location = New-Object System.Drawing.Point(0,0)


$labelRevokeFrom = New-Object System.Windows.Forms.Label
$labelRevokeFrom.Text = "From user (User's calendar)"
$labelRevokeFrom.width = 200
$labelRevokeFrom.height = 20
$labelRevokeFrom.Location = New-Object System.Drawing.Point(130,100)

$InputRevokeFrom = New-Object System.Windows.Forms.TextBox
$InputRevokeFrom.height = 10
$InputRevokeFrom.width = 100
$InputRevokeFrom.Location = New-Object System.Drawing.Point(150,125)

$labelRevokeTo = New-Object System.Windows.Forms.Label 
$labelRevokeTo.Text = "To user (User losing permission)"
$labelRevokeTo.width = 200
$labelRevokeTo.height = 20
$labelRevokeTo.Location = New-Object System.Drawing.Point(130,155)

$InputRevokeTo = New-Object System.Windows.Forms.TextBox
$InputRevokeTo.height = 10
$InputRevokeTo.width = 100
$InputRevokeTo.Location = New-Object System.Drawing.Point(150,175)

$ButtonRevokeCalendar = New-Object System.Windows.Forms.Button 
$ButtonRevokeCalendar.text = "Revoke Calendar"
$ButtonRevokeCalendar.width = 75
$ButtonRevokeCalendar.height = 50 
$ButtonRevokeCalendar.Location = New-Object System.Drawing.Point(150,200)
$ButtonRevokeCalendar.Add_Click({
    RevokeCalendar
})

$ButtonReturnRevoke = New-Object System.Windows.Forms.Button 
$ButtonReturnRevoke.text = "Return to menu"
$ButtonReturnRevoke.width = 75
$ButtonReturnRevoke.height = 50
$ButtonReturnRevoke.Location = New-Object System.Drawing.Point(100,300)
$ButtonReturnRevoke.Add_Click({
    SystemReturn
})

$ButtonQuitRevoke = New-Object System.Windows.Forms.Button
$ButtonQuitRevoke.Text = "Close system"
$ButtonQuitRevoke.Width = 60
$ButtonQuitRevoke.Height = 50
$ButtonQuitRevoke.Location = New-Object System.Drawing.Point(200,300)
$ButtonQuitRevoke.Add_Click({
    SystemQuit
})

# Functions

Function SystemModuleInstall {
    Install-Module -Name ExchangeOnlineManagement
}

Function SystemLogin {
    $liveCred = Get-Credential
    $Credentials = $LiveCred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication Basic -Credential $LiveCred
        Import-PSSession $Session
    $Credentials
}

Function SystemQuit {
    $Form.Close()
}

Function SystemReturn {
    $GroupboxMenu.Visible = $true
    $GroupboxShare.Visible = $false
    $GroupboxGet.Visible = $false 
    $GroupboxRevoke.Visible = $false 
}
Function SystemLoadMenu {
    $GroupboxLogin.Visible = $false
    $GroupboxMenu.Visible = $true 
}

Function SystemHideMenu {
    $GroupboxMenu.Visible = $false
}

Function SystemShareCalendarMenu {
    $GroupboxShare.Visible = $true
}

Function SystemGetCalendarMenu {
    $GroupboxGet.Visible = $true 
}

Function SystemRevokeCalendarMenu {
    $GroupboxRevoke.Visible = $true
}

Function ShareCalendar {
    $FromUser = $InputShareFrom.Text 
    $ToUser =   $InputShareTo.Text 
    Add-MailboxFolderPermission -Identity ${FromUser}:\Calendar -USer ${ToUser} -AccessRights Editor 
    [System.Windows.MessageBox]::Show('Complete')
}

Function GetCalendarPermissions {
    $CalendarInput = $InputGetPermissions.Text
    $CalendarPath = "C:\temp\CalendarPermissions.csv"
    Get-MailboxFolderPermission -Identity ${CalendarInput}:\Calendar | Export-Csv -Path $CalendarPath -NoTypeInformation 
    [System.Windows.MessageBox]::Show('Complete')
    Start-Process explorer $CalendarPath
}

Function RevokeCalendar {
    $CalendarToRevoke = $InputRevokeFrom.Text
    $CalendarFromRevoke = $InputRevokeTo.Text
    Remove-MailboxFolderPermission -Identity ${CalendarToRevoke}:\Calendar -User ${CalendarFromRevoke}
    [System.Windows.MessageBox]::Show('Complete')
}

$GroupboxLogin.Controls.AddRange(@($ButtonInstall, $ButtonLogin, $ButtonQuit))
$GroupboxMenu.Controls.AddRange(@($ButtonGrant, $ButtonGet, $ButtonRevoke))
$GroupboxShare.Controls.AddRange(@($labelShareFrom, $InputShareFrom, $InputShareTo, $labelShareTo, $ButtonShareCalendar, $ButtonReturnShare, $ButtonQuitShare))
$GroupboxGet.Controls.AddRange(@($InputGetPermissions, $ButtonGetPermissions, $ButtonReturnGet, $ButtonQuitGet))
$GroupboxRevoke.Controls.AddRange(@($labelRevokeFrom, $InputRevokeFrom, $labelRevoketo, $InputRevokeTo, $ButtonRevokeCalendar, $ButtonReturnRevoke, $ButtonQuitRevoke))

$Form.Controls.AddRange(@($GroupboxLogin, $GroupboxMenu, $GroupboxShare, $GroupboxGet, $GroupboxRevoke))
$Form.ShowDialog()