Add-Type -AssemblyName System.Windows.Forms #Form描画のため
Add-Type -AssemblyName System.Drawing #Form描画のため
#Add-Type -AssemblyName System.Net.NetworkInformation #Networkの情報取得のため(ver2.0だと読み出せない)
Add-Type -AssemblyName System.Net #Networkの情報取得のため(ver2.0だとこちらは読み出し可能)

# -------------------------------------------------------------
# 環境
# -------------------------------------------------------------
$StateDataTable = [xml](Get-Content ".\test.xml")

# -------------------------------------------------------------
# インターフェース選択フォーム
# -------------------------------------------------------------
# フォーム全体の設定
#$Form1 = [System.Windows.Forms.Form]@{ #ver3.0だとこう書ける
$form1 = New-Object System.Windows.Forms.Form -Property @{
    Text = "IPアドレス変更ツール"
    Size = New-Object System.Drawing.Size(300,420)
    StartPosition = "CenterScreen"
    #Topmost = $True  # フォームを最前面に表示
}

# ラベル(NIC選択)の設定
$Niclabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,10)
    Size = New-Object System.Drawing.Size(230,20)
    Text = "使用するネットワークを選択"
}
# コンボボックス(NIC選択)の設定
$NicComboBox = New-Object System.Windows.Forms.ComboBox -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(230,20)
    DropDownStyle = "DropDown"
}
# ラベル(NIC選択)の設定
$Statelabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,60)
    Size = New-Object System.Drawing.Size(150,20)
    Text = "適用する設定を選択"
}
# リストボックス(NIC選択)の設定
$StateListBox = New-Object System.Windows.Forms.ListBox -Property @{
    Location = New-Object System.Drawing.Point(10,80)
    Size = New-Object System.Drawing.Size(150,80)
}
# 変更ボタンの設定
$ChangeButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(180,80)
    Size = New-Object System.Drawing.Size(80,30)
    Text = "設定変更"
    Enabled = $False
    #DialogResult = [System.Windows.Forms.DialogResult]::OK
}
# 現行設定確認ボタンの設定
$CurrentButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(180,120)
    Size = New-Object System.Drawing.Size(80,30)
    Text = "現行確認"
    Enabled = $False
    #DialogResult = [System.Windows.Forms.DialogResult]::OK
}
# ラベル(選択したリストの詳細)の設定
$Selectedlabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,170)
    Size = New-Object System.Drawing.Size(230,20)
    Text = "選択した設定の詳細"
}
# リストビュー(選択したリストの詳細)の設定
$SelectedIPView = New-Object System.Windows.Forms.ListView -Property @{
    Location = New-Object System.Drawing.Point(10,190)
    Size = New-Object System.Drawing.Size(230,80)
    View = "Details"
    GridLines = $False
}
# リストビュー(選択したリストの詳細)の設定
$SelectedGWView = New-Object System.Windows.Forms.ListView -Property @{
    Location = New-Object System.Drawing.Point(10,280)
    Size = New-Object System.Drawing.Size(230,80)
    View = "Details"
    GridLines = $False
}

# フォーム全体の設定
#$Form1 = [System.Windows.Forms.Form]@{ #ver3.0だとこう書ける
$form2 = New-Object System.Windows.Forms.Form -Property @{
    Text = "IPアドレス変更ツール"
    Size = New-Object System.Drawing.Size(520,400)
    StartPosition = "CenterScreen"
    #Topmost = $True  # フォームを最前面に表示
}

# ラベル(選択NICの今の設定)の設定
$Currentlabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,10)
    Size = New-Object System.Drawing.Size(480,20)
    Text = "現在の設定"
}
# テキストボックス(選択NICの今の設定)の設定
$CurrentText = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(480,330)
    Text = ""
    Font = new-object System.Drawing.Font("MS Pゴシック", 10)
    Multiline = $True
    ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    WordWrap = $False;
    ReadOnly = $True
}

# フォームにコントロールを追加
$form1.Controls.Add($NicLabel)
$form1.Controls.Add($NicComboBox)
$form1.Controls.Add($StateLabel)
$form1.Controls.Add($StateListBox)
$form1.Controls.Add($ChangeButton)
$form1.Controls.Add($CurrentButton)
$form1.Controls.Add($SelectedLabel)
$form1.Controls.Add($SelectedIPView)
$form1.Controls.Add($SelectedGWView)

# フォームにコントロールを追加
$form2.Controls.Add($CurrentLabel)
$form2.Controls.Add($CurrentText)

# リストビューにヘッダーを追加
[void]$SelectedIPView.Columns.Add("アドレス",110)
[void]$SelectedIPView.Columns.Add("サブネット",110)
[void]$SelectedGWView.Columns.Add("デフォルトゲートウェイ",160)
[void]$SelectedGWView.Columns.Add("メトリック",60)

# コンボボックスに要素追加
function NicComboBox_AddItem {
    [System.Net.NetworkInformation.NetworkInterface]::GetAllNetWorkInterfaces() `
    | ForEach { $_name=$_.Name; $_.GetIpProperties() } `
    | Where { $_.UnicastAddresses.Count -gt 0 } `
    | ForEach {
        Write-Host $_name
        [void]$NicComboBox.Items.Add($_name)
    }
}

# リストボックスに要素追加の設定
function StateListBox_AddItem {
    $StateDataTable `
    | Select -ExpandProperty Data `
    | Select -ExpandProperty IPAddressSet `
    | ForEach {
        Write-Host $_.Name
        [void]$StateListBox.Items.Add($_.Name)
    }
}

# リスト選択時のテキストボックスにIPアドレスを表示させるイベントハンドラーの定義
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

# リスト選択時のリストビュにIPアドレスを表示させるイベントハンドラーの定義
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
        [System.Windows.Forms.MessageBox]::Show("インターフェースを選択してください", "エラー", "OK", "Information")
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
