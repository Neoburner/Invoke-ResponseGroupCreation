###########################
# Current version: 0.1
#
# Change Notes:
# 10/08/2017 - Basic Script  Created and Tested

###########################
# Requirements:
# Lync Powershell Module
# 

# Global Varibles
$parentService = ""

# DO NOT EDIT BELOW


###########################
# Requirements:
# Lync Powershell Module
#

# Menu Function

Function ShowMenu {

    Write-Host " "
    Write-Host "*********************************"
    Write-Host "KCC Response Group Creation v0.1"
    Write-Host "*********************************"
    Write-Host " "
    Write-Host "1: Create Groups"
    Write-Host "2: Create Queues"
    Write-Host "3: Create Workflows"
    Write-Host " "

}

Function Invoke-CreateGroups () {
    
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    $Users = Import-Csv -Path $inputfile
                ForEach($User in $Users){
                    
                    $groupName = $($User.groupName)
                    $groupDescription = $($User.groupDescription)
                    $participationPolicy = $($User.participationPolicy)
                    $alertTime = $($User.alertTime)
                    $routingMethod = $($User.routingMethod)
                    
                    Start-Sleep (3);
                    New-CsRgsAgentGroup -Parent $parentService -Name $groupName -Description $groupDescription -ParticipationPolicy $participationPolicy -AgentAlertTime $alertTime -RoutingMethod $routingMethod
                    Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "Group Created"
                    Start-Sleep (2);
                }
               
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "$groupName - Group creation Complete..."
    Start-Sleep (2);
    ShowMenu
}

Function Invoke-CreateQueues () {
    
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    $Users = Import-Csv -Path $inputfile
                ForEach($User in $Users){
                    
                    $groupName = $($User.groupName)
                    $groupDescription = $($User.groupDescription)
                    $queueVoicemail = $($User.queueVoicemail)
                    $queueTimeout = $($User.queueTimeout)
                    $groupIdentity = (Get-CsRgsAgentGroup -Name $groupName).Identity
                    
                    Start-Sleep (3);
                    
                    If (!$queueVoicemail) {
                    New-CsRgsQueue -Parent $parentService -Name $groupName -Description $groupDescription -AgentGroupIDList $groupIdentity -TimeoutThreshold $queueTimeout
                    
                    Write-Host -ForegroundColor Cyan -BackgroundColor Red $groupName "Queue Created"
                    Start-Sleep (2);    
                    }
                    
                    Else {
                    $x = New-CsRgsCallAction -Action TransferToVoicemailUri -Uri $queueVoicemail
                    New-CsRgsQueue -Parent $parentService -Name $groupName -Description $groupDescription -AgentGroupIDList $groupIdentity -TimeoutThreshold $queueTimeout -TimeoutAction $x
                    
                    Write-Host -ForegroundColor Cyan -BackgroundColor Red $groupName "Queue Created"
                    Start-Sleep (2);  
                    }
                    
                }
               
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "All Queues created..."
    Start-Sleep (2);
    ShowMenu
}


Function Invoke-CreateWorkflow () {
    
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    $Users = Import-Csv -Path $inputfile
                ForEach($User in $Users){
                    
                    $groupName = $($User.groupName)
                    $groupDescription = $($User.groupDescription)
                    $queueIdentity = (Get-CsRgsQueue -Identity $parentService -Name $groupName).Identity
                    $workflowURI = $($User.workflowURI)
                    
                    Start-Sleep (3);
                    
                    $callAction = New-CsRgsCallAction -Action TransferToQueue -QueueId $queueIdentity
                    New-CsRgsWorkflow -Parent $parentService -Name $groupName -Description $groupDescription -PrimaryUri $workflowURI -DefaultAction $callAction -Active $true
                    
                    Write-Host -ForegroundColor Cyan -BackgroundColor Red $groupName "Workflow Created"
                    Start-Sleep (2);  
                    }
                    
                
               
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "All Workflows created..."
    Start-Sleep (2);
    ShowMenu
}
# Execute Menu
do
 {
     ShowMenu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
           '1' {
             Invoke-CreateGroups
         } '2' {
             Invoke-CreateQueues
         } '3' {
             Invoke-CreateWorkflow
         }
     }
     pause
 }
 until ($selection -eq 'q')