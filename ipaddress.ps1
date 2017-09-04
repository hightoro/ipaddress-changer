Add-Type -AssemblyName System.Windows.Forms #Form�`��̂���
Add-Type -AssemblyName System.Drawing #Form�`��̂���
#Add-Type -AssemblyName System.Net.NetworkInformation #Network�̏��擾�̂���(ver2.0���Ɠǂݏo���Ȃ�)
Add-Type -AssemblyName System.Net #Network�̏��擾�̂���(ver2.0���Ƃ�����͓ǂݏo���\)

# -------------------------------------------------------------
# ��
# -------------------------------------------------------------
$StateDataTable = [xml](Get-Content ".\test.xml")

# -------------------------------------------------------------
# �C���^�[�t�F�[�X�I���t�H�[��
# -------------------------------------------------------------
# �t�H�[���S�̂̐ݒ�
#$Form1 = [System.Windows.Forms.Form]@{ #ver3.0���Ƃ���������
$form1 = New-Object System.Windows.Forms.Form -Property @{
    Text = "IP�A�h���X�ύX�c�[��"
    Size = New-Object System.Drawing.Size(300,420)
    StartPosition = "CenterScreen"
    #Topmost = $True  # �t�H�[�����őO�ʂɕ\��
}

# ���x��(NIC�I��)�̐ݒ�
$Niclabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,10)
    Size = New-Object System.Drawing.Size(230,20)
    Text = "�g�p����l�b�g���[�N��I��"
}
# �R���{�{�b�N�X(NIC�I��)�̐ݒ�
$NicComboBox = New-Object System.Windows.Forms.ComboBox -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(230,20)
    DropDownStyle = "DropDown"
}
# ���x��(NIC�I��)�̐ݒ�
$Statelabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,60)
    Size = New-Object System.Drawing.Size(150,20)
    Text = "�K�p����ݒ��I��"
}
# ���X�g�{�b�N�X(NIC�I��)�̐ݒ�
$StateListBox = New-Object System.Windows.Forms.ListBox -Property @{
    Location = New-Object System.Drawing.Point(10,80)
    Size = New-Object System.Drawing.Size(150,80)
}
# �ύX�{�^���̐ݒ�
$ChangeButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(180,80)
    Size = New-Object System.Drawing.Size(80,30)
    Text = "�ݒ�ύX"
    Enabled = $False
    #DialogResult = [System.Windows.Forms.DialogResult]::OK
}
# ���s�ݒ�m�F�{�^���̐ݒ�
$CurrentButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(180,120)
    Size = New-Object System.Drawing.Size(80,30)
    Text = "���s�m�F"
    Enabled = $False
    #DialogResult = [System.Windows.Forms.DialogResult]::OK
}
# ���x��(�I���������X�g�̏ڍ�)�̐ݒ�
$Selectedlabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,170)
    Size = New-Object System.Drawing.Size(230,20)
    Text = "�I�������ݒ�̏ڍ�"
}
# ���X�g�r���[(�I���������X�g�̏ڍ�)�̐ݒ�
$SelectedIPView = New-Object System.Windows.Forms.ListView -Property @{
    Location = New-Object System.Drawing.Point(10,190)
    Size = New-Object System.Drawing.Size(230,80)
    View = "Details"
    GridLines = $False
}
# ���X�g�r���[(�I���������X�g�̏ڍ�)�̐ݒ�
$SelectedGWView = New-Object System.Windows.Forms.ListView -Property @{
    Location = New-Object System.Drawing.Point(10,280)
    Size = New-Object System.Drawing.Size(230,80)
    View = "Details"
    GridLines = $False
}

# �t�H�[���S�̂̐ݒ�
#$Form1 = [System.Windows.Forms.Form]@{ #ver3.0���Ƃ���������
$form2 = New-Object System.Windows.Forms.Form -Property @{
    Text = "IP�A�h���X�ύX�c�[��"
    Size = New-Object System.Drawing.Size(520,400)
    StartPosition = "CenterScreen"
    #Topmost = $True  # �t�H�[�����őO�ʂɕ\��
}

# ���x��(�I��NIC�̍��̐ݒ�)�̐ݒ�
$Currentlabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,10)
    Size = New-Object System.Drawing.Size(480,20)
    Text = "���݂̐ݒ�"
}
# �e�L�X�g�{�b�N�X(�I��NIC�̍��̐ݒ�)�̐ݒ�
$CurrentText = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(480,330)
    Text = ""
    Font = new-object System.Drawing.Font("MS P�S�V�b�N", 10)
    Multiline = $True
    ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    WordWrap = $False;
    ReadOnly = $True
}

# �t�H�[���ɃR���g���[����ǉ�
$form1.Controls.Add($NicLabel)
$form1.Controls.Add($NicComboBox)
$form1.Controls.Add($StateLabel)
$form1.Controls.Add($StateListBox)
$form1.Controls.Add($ChangeButton)
$form1.Controls.Add($CurrentButton)
$form1.Controls.Add($SelectedLabel)
$form1.Controls.Add($SelectedIPView)
$form1.Controls.Add($SelectedGWView)

# �t�H�[���ɃR���g���[����ǉ�
$form2.Controls.Add($CurrentLabel)
$form2.Controls.Add($CurrentText)

# ���X�g�r���[�Ƀw�b�_�[��ǉ�
[void]$SelectedIPView.Columns.Add("�A�h���X",110)
[void]$SelectedIPView.Columns.Add("�T�u�l�b�g",110)
[void]$SelectedGWView.Columns.Add("�f�t�H���g�Q�[�g�E�F�C",160)
[void]$SelectedGWView.Columns.Add("���g���b�N",60)

# �R���{�{�b�N�X�ɗv�f�ǉ�
function NicComboBox_AddItem {
    [System.Net.NetworkInformation.NetworkInterface]::GetAllNetWorkInterfaces() `
    | ForEach { $_name=$_.Name; $_.GetIpProperties() } `
    | Where { $_.UnicastAddresses.Count -gt 0 } `
    | ForEach {
        Write-Host $_name
        [void]$NicComboBox.Items.Add($_name)
    }
}

# ���X�g�{�b�N�X�ɗv�f�ǉ��̐ݒ�
function StateListBox_AddItem {
    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | ForEach {
        Write-Host $_.Name
        [void]$StateListBox.Items.Add($_.Name)
    }
}

# ���X�g�I�����̃e�L�X�g�{�b�N�X��IP�A�h���X��\��������C�x���g�n���h���[�̒�`
$NicComboBox.Add_SelectedIndexChanged({
    $CurrentButton.Enabled = $true
})
$CurrentButton.Add_Click({

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "netsh"
    #$pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    #$pinfo.Arguments = ("interface ip show address '{0}'" -f $NicComboBox.SelectedItem)
    $pinfo.Arguments = ("interface ip show address `"{0}`"" -f $NicComboBox.SelectedItem)
    $pinfo.WindowStyle = "Hidden"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    #$stdout = $p.StandardOutput.ReadToEnd()

    #Invoke-Expression ("netsh interface ip show address '{0}'" -f $NicComboBox.SelectedItem)
    #$t = (Invoke-Expression ("netsh interface ip show address '{0}'" -f $NicComboBox.SelectedItem) ) -Replace "`r`n","`n"
    $CurrentText.Text = $p.StandardOutput.ReadToEnd()

    $Form2.ShowDialog()
})

# ���X�g�I�����̃��X�g�r����IP�A�h���X��\��������C�x���g�n���h���[�̒�`
$StateListBox.Add_SelectedIndexChanged({
    [void]$SelectedIPView.Items.Clear()
    [void]$SelectedGWView.Items.Clear()

    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | Where {$_.Name -eq $StateListBox.SelectedItem} `
    | Select -ExpandProperty IP `
    | ForEach {
        $Item =  New-Object System.Windows.Forms.ListViewItem( $_.Address )
        [void]$Item.SubItems.Add( $_.Subnet )
        [void]$SelectedIPView.Items.Add($Item)
    }

    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | Where {$_.Name -eq $StateListBox.SelectedItem} `
    | Select -ExpandProperty GW `
    | ForEach {
        $Item =  New-Object System.Windows.Forms.ListViewItem( $_.Address )
        [void]$Item.SubItems.Add( $_.Metric )
        [void]$SelectedGWView.Items.Add($Item)
    }
    $ChangeButton.Enabled = $true
})

$ChangeButton.Add_Click({
    if( $NicComboBox.SelectedItem -eq $null){
        [System.Windows.Forms.MessageBox]::Show("�C���^�[�t�F�[�X��I�����Ă�������", "�G���[", "OK", "Information")
        return
    }

    $cnt = 0
    $list = "-Command "

    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | Where {$_.Name -eq $StateListBox.SelectedItem} `
    | Select -ExpandProperty IP `
    | ForEach {
        if($cnt -eq 0){
            $list += ("netsh interface ip set address '{0}' static {1} {2};" -f $NicComboBox.SelectedItem,$_.Address,$_.Subnet)
        } else {
            $list += ("netsh interface ip add address name='{0}' addr={1} mask={2};" -f $NicComboBox.SelectedItem,$_.Address,$_.Subnet)
        }
        $cnt++
    }

    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | Where {$_.Name -eq $StateListBox.SelectedItem} `
    | Select -ExpandProperty GW `
    | ForEach {
        $list += ("netsh interface ip add address name='{0}' gateway={1} gwmetric={2};" -f $NicComboBox.SelectedItem,$_.Address,$_.Metric)
    }
    Write-host $list
    Start-Process powershell -ArgumentList $list -Verb RunAs -Wait
})


NicComboBox_AddItem
StateListBox_AddItem

$Form1.ShowDialog()
