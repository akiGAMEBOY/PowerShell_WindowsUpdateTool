#################################################################################
# 処理名　 | WindowsUpdateTool
# 機能　　 | 辞書ファイルの単語検索ツール
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | なし
#################################################################################
# 設定
# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み
#   フォーム用
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = 'Stop'
# 定数
[System.String]$c_config_file = 'setup.ini'
# エラーコード enum設定
Add-Type -TypeDefinition @"
    public enum MESSAGECODE {
        Successful = 0,
        Abend,
        Cancel,
        Info_LoadedSettingfile,
        Info_SkipUpdate,
        Info_SkipMSDefenderInvalid,
        Info_SkipSelectWindowsUpdate,
        Info_SkipSelectMSDefender,
        Info_SkipSelectWinget,
        Info_SkipExcludeWinget,
        Confirm_ExecutionTool,
        Confirm_ExecuteWindowsUpdate,
        Confirm_ExecuteMSDedender,
        Confirm_ExecuteWinget,
        Confirm_ExecuteWinget_Individual,
        Error_NotAdmin,
        Error_LoadingSettingfile,
        Error_EmptyTargetfolder,
        Error_NotCheckbox,
        Error_MaxRetries,
        Error_InstallModules,
        Error_GetWinUpdate,
        Error_WinUpdate_all,
        Error_WinUpdate_Individual,
        Error_MSDefenderStatusCheck,
        Error_UpdateMSDefender_Ex,
        Error_UpdateMSDefender_Returnerror,
        Error_UpdateWinget_Ex,
        Error_UpdateWinget_Individual_Ex,
        Error_UpdateWinget_Returnerror,
        Error_CheckWinget,
        Error_WingetNotInstall,
        Error_WingetUpgrade
    }
"@

### DEBUG ###
Set-Variable -Name "DEBUG_ON" -Value $false -Option Constant

### Function --- 開始 --->
#################################################################################
# 処理名　 | RemoveDoubleQuotes
# 機能　　 | 先頭桁と最終桁にあるダブルクォーテーションを削除
#--------------------------------------------------------------------------------
# 戻り値　 | String（削除後の文字列）
# 引数　　 | target_str: 対象文字列
#################################################################################
Function RemoveDoubleQuotes {
    Param (
        [System.String]$target_str
    )

    [System.String]$removed_str = $target_str
    
    if ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq '"') -and
            ($target_str.Substring($target_str.Length - 1, 1) -eq '"')) {
            # 先頭桁と最終桁のダブルクォーテーション削除
            $removed_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''
        
        Write-Host "Function RemoveDoubleQuotes: target_str  [${target_str}]"
        Write-Host "                             removed_str [${removed_str}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $removed_str
}

#################################################################################
# 処理名　 | isAdminPowerShell
# 機能　　 | PowerShellが管理者として実行しているか確認
#          | 参考情報：https://zenn.dev/haretokidoki/articles/67788ca9b47b27
#--------------------------------------------------------------------------------
# 戻り値　 | Boolean（True: 管理者権限あり, False: 管理者権限なし）
# 引数　　 | -
#################################################################################
Function isAdminPowerShell {
    $win_id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $win_principal = new-object System.Security.Principal.WindowsPrincipal($win_id)
    $admin_permission = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    return $win_principal.IsInRole($admin_permission)
}

#################################################################################
# 処理名　 | GetPsCharcode
# 機能　　 | PowerShellコンソールの文字コードを取得
#          | 参考情報：https://zenn.dev/haretokidoki/articles/67788ca9b47b27
#--------------------------------------------------------------------------------
# 戻り値　 | ps_charcode[]
#          |  - 項目01 文字エンコードを指定できるコマンドレットの既定値
#          |  - 項目02 PowerShellから外部プログラムに渡す文字エンコードの設定
#          |  - 項目01 PowerShellのコンソールに出力する文字エンコードの設定
# 引数　　 | -
#################################################################################
Function GetPsCharcode {
    [System.String[]]$ps_charcode = @()
    $ps_charcode = @(
        # 文字エンコードを指定できるコマンドレットの既定値
        ($PSDefaultParameterValues['*:Encoding']),
        # PowerShellから外部プログラムに渡す文字エンコードの設定
        ($global:OutputEncoding).WebName,
        # PowerShellのコンソールに出力する文字エンコードの設定
        ([console]::OutputEncoding).WebName
    )

    return $ps_charcode
}

#################################################################################
# 処理名　 | ChangeWindowTitle
# 機能　　 | PowerShellウィンドウのタイトル変更（文字コードとPowerShellの管理者権限有無を追加）
#          | 参考情報：https://zenn.dev/haretokidoki/articles/67788ca9b47b27
#--------------------------------------------------------------------------------
# 戻り値　 | -
# 引数　　 | -
#################################################################################
# PowerShellウィンドウのタイトル変更
Function ChangeWindowTitle {
    # 区切り文字の設定
    [System.String]$pos1 = '|'
    [System.String]$pos2 = ';'

    # 現在のタイトルを取得
    [System.String]$title = $Host.UI.RawUI.WindowTitle
    [System.String]$base_title = $title

    # 既にこのFunctionでタイトル変更している場合、一番左にある文字列を抽出
    [System.String[]]$title_array = $title.Split($pos1)
    if ($title_array.Length -ne 0) {
        $base_title = ($title_array[0]).TrimEnd()
    }

    # 現在の文字コードを取得しタイトルに追加
    [System.String[]]$ps_charcode = GetPsCharcode

    [System.String]$change_title = $base_title
    if (isAdminPowerShell) {
        # 管理者として実行している場合
        $change_title = $base_title + " $pos1 " +
                        "DefaultParameter='$($ps_charcode[0])'" + " $pos2 " +
                        "GlobalEncoding='$($ps_charcode[1])'" + " $pos2 " +
                        "ConsoleEncoding='$($ps_charcode[2])'" + " $pos2 " +
                        "#Administrator"
    }
    else {
        # 管理者として実行していない場合
        $change_title = $base_title + " $pos1 " +
                        "DefaultParameter='$($ps_charcode[0])'" + " $pos2 " +
                        "GlobalEncoding='$($ps_charcode[1])'" + " $pos2 " +
                        "ConsoleEncoding='$($ps_charcode[2])'" + " $pos2 " +
                        "#Not_Administrator"
    }
    $Host.UI.RawUI.WindowTitle = $change_title

    # 完了メッセージ
    Write-Host 'タイトルに“文字コード”と“管理者権限の有無”の情報を追加しました。' -ForegroundColor Cyan
    Write-Host ''
    Write-Host ''
}

#################################################################################
# 処理名　 | SetPsOutputEncoding
# 機能　　 | PowerShellにおける複数の文字コード設定を一括変更
#          | 参考情報：https://zenn.dev/haretokidoki/articles/8946231076f129
#--------------------------------------------------------------------------------
# 戻り値　 | -
# 引数　　 | $charcode（引数を省略した場合は、'reset_encoding'で設定）
# 　　　　 |  - 'utf8'          : UTF-8に設定
# 　　　　 |  - 'sjis'          : Shift-JISに設定
# 　　　　 |  - 'ascii'         : US-ASCIIに設定
# 　　　　 |  - 'rm_encoding'   : デフォルトパラーメーターを削除
# 　　　　 |  - 'reset_encoding': 規定値に戻す
#################################################################################
Function SetPsOutputEncoding {
    Param (
        [System.String]$charcode = 'reset_encoding'
    )

    switch ($charcode) {
        # 文字エンコードをUTF8に設定する
        'utf8' {
            $PSDefaultParameterValues['*:Encoding'] = 'utf8'
            $global:OutputEncoding = [System.Text.Encoding]::UTF8
            [console]::OutputEncoding = [System.Text.Encoding]::UTF8
        }
        # 文字エンコードをShift JIS（SJIS）に設定する
        'sjis' {
            # $PSDefaultParameterValues['*:Encoding'] = 'default'について
            #   この設定はCore以外（5.1以前）の環境でのみShift JISで設定される。
            #   Core環境のデフォルト値は、UTF-8でありUTF-8で設定されてしまう。
            #   また、Shift JISのパラメーターも存在しない為、Core環境でShift JISの設定は不可となる。
            $PSDefaultParameterValues['*:Encoding'] = 'default'
            $global:OutputEncoding = [System.Text.Encoding]::GetEncoding('shift_jis')
            [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding('shift_jis')
        }
        # 文字エンコードをASCIIに設定する
        'ascii' {
            $PSDefaultParameterValues.Remove('*:Encoding')
            $global:OutputEncoding = [System.Text.Encoding]::ASCII
            [console]::OutputEncoding = [System.Text.Encoding]::ASCII
        }
        # デフォルトパラメータの文字エンコード指定を解除する
        'rm_encoding' {
            $PSDefaultParameterValues.Remove('*:Encoding')
        }
        # 文字エンコード設定を初期状態に戻す
        'reset_encoding' {
            $PSDefaultParameterValues.Remove('*:Encoding')

            if ($PSVersionTable.PSEdition -eq 'Core') {
                # Core の場合
                $global:OutputEncoding = [System.Text.Encoding]::UTF8
                [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding('shift_jis')
            }
            else {
                # Core 以外の場合（PowerShell 5.1 以前）
                $global:OutputEncoding = [System.Text.Encoding]::ASCII
                [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding('shift_jis')
            }
        }
    }
    # タイトルの表示切替Function呼び出し
    ChangeWindowTitle
}

#################################################################################
# 処理名　 | AcquisitionFormsize
# 機能　　 | Windowsフォーム用のサイズをモニターサイズから除算で設定
#--------------------------------------------------------------------------------
# 戻り値　 | String[]（変換後のサイズ：1要素目 横サイズ、2要素目 縦サイズ）
# 引数　　 | divisor: 除数（モニターサイズから除算するため）
#################################################################################
Function AcquisitionFormsize {
    Param (
        [System.UInt32]$divisor
    )

    # 現在のモニターサイズを取得
    [Microsoft.Management.Infrastructure.CimInstance]$graphics_info = (Get-CimInstance -ClassName Win32_VideoController)
    [System.UInt32]$width = $graphics_info.CurrentHorizontalResolution
    [System.UInt32]$height = $graphics_info.CurrentVerticalResolution

    # モニターのサイズから除数で割る
    [System.UInt32]$form_width = $width / $divisor
    [System.UInt32]$form_height = $height / $divisor
    
    [System.UInt32[]]$form_size = @($form_width, $form_height)

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function AcquisitionFormsize: form_size [${form_size}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $form_size
}

#################################################################################
# 処理名　 | ConfirmYesno
# 機能　　 | YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　 | Boolean（True: 正常終了, False: 処理中断）
# 引数　　 | prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno {
    Param (
        [System.String]$prompt_message,
        [System.String]$prompt_title='実行前の確認'
    )

    # 除数「5」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(5)

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(($form_size[0] * 0.016), ($form_size[1] * 0.030))
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(($form_size[0] * 0.0156),($form_size[1] * 0.0285))
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.07))
    # $label.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.3))
    $label.Text = $prompt_message
    $label.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = 'OK'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # フォーム表示
    [System.Boolean]$is_selected = ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: is_selected [${is_selected}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $is_selected
}

#################################################################################
# 処理名　 | ValidateInputValues
# 機能　　 | 入力値の検証
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | setting_parameters[]
# 　　　　 |  - 項目01 Windows Update を更新する
# 　　　　 |  - 項目02 Microsoft Defender を更新する
# 　　　　 |  - 項目03 アプリ を更新する
#################################################################################
Function ValidateInputValues {
    Param (
        [System.Boolean[]]$setting_parameters
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful

    # メッセージボックス用
    [System.String]$messagebox_title = ''
    [System.String]$messagebox_messages = ''

    # チェックボックス
    # Windows Update / Microsoft Defnerder / アプリ いずれもチェックされていない場合
    if ((-Not($setting_parameters[0])) -and
        (-Not($setting_parameters[1])) -and
        (-Not($setting_parameters[2]))) {
        $messagecode = [MESSAGECODE]::Error_NotCheckbox
        $messagebox_messages = RetrieveMessage $messagecode
        $messagebox_title = '入力チェック'
        ShowMessagebox $messagebox_messages $messagebox_title
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: return [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　 | SwitchActiveWindow
# 機能　　 | アクティブウィンドウの切り替え
#--------------------------------------------------------------------------------
# 戻り値　 | なし
# 引数　　 | cscode: アクティブウィンドウ切り替えるC#のコード
# 　　　　 | window_name: 切り替えるウィンドウの名前
#################################################################################
Function SwitchActiveWindow {
    Param (
        [System.String]$cscode_filepath,
        [System.String]$window_name
    )

    add-type -AssemblyName microsoft.VisualBasic
    add-type -AssemblyName System.Windows.Forms

    [System.String]$cscode = Get-Content $cscode_filepath -Raw -Encoding utf8

    $Win32 = add-type -memberDefinition $cscode -name "Win32ApiFunctions" -passthru
    
    $ps = Get-Process | Where-Object {$_.Name -match $window_name}
    foreach($process in $ps){
        $Win32::ActiveWindow($process.MainWindowHandle);
    }
}

#################################################################################
# 処理名　 | SettingInputValues
# 機能　　 | 入力フォルダーの設定（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　 | Object[]
# 　　　　 |  - 項目01 対象フォルダー        : 画面での設定値 - ツールの作業フォルダーとして使用
# 引数　　 | function_parameters[]
# 　　　　 |  - 項目01 ツール実行場所                : ツールの実行場所
# 　　　　 |  - 項目02 Windows Update を更新する     : 初期表示用の値 - 設定ファイルの設定値が反映
# 　　　　 |  - 項目03 Microsoft Defender を更新する : 初期表示用の値 - 設定ファイルの設定値が反映
# 　　　　 |  - 項目04 アプリ を更新する             : 初期表示用の値 - 設定ファイルの設定値が反映
# 　　　　 |  - 項目05 一括 or 個別 インストール     : 初期表示用の値 - 設定ファイルの設定値が反映
#################################################################################
Function SettingInputValues {
    Param (
        [System.Object[]]$function_parameters
    )

    # 除数「3」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(3)

    # フォームの作成
    [System.String]$prompt_title = '実行前の設定'
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$($function_parameters[0])\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # Windows Update を更新する - チェックボックスの作成
    [System.Windows.Forms.CheckBox]$checkbox_selected_windowsupdate = New-Object System.Windows.Forms.CheckBox
    $checkbox_selected_windowsupdate.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.070))
    $checkbox_selected_windowsupdate.Size = New-Object System.Drawing.Size(20, 20)
    $checkbox_selected_windowsupdate.Checked = [System.Convert]::ToBoolean($function_parameters[1])
    $checkbox_selected_windowsupdate.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # Windows Update を更新する - ラベル作成
    [System.Windows.Forms.Label]$label_selected_windowsupdate = New-Object System.Windows.Forms.Label
    $label_selected_windowsupdate.Location = New-Object System.Drawing.Point(($form_size[0] *0.08),($form_size[1] * 0.070))
    $label_selected_windowsupdate.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_selected_windowsupdate.Text = 'Windows Update の更新を実施する'
    $label_selected_windowsupdate.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # Microsoft Defender を更新する - チェックボックスの作成
    [System.Windows.Forms.CheckBox]$checkbox_selected_msdefender = New-Object System.Windows.Forms.CheckBox
    $checkbox_selected_msdefender.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.175))
    $checkbox_selected_msdefender.Size = New-Object System.Drawing.Size(20, 20)
    $checkbox_selected_msdefender.Checked = [System.Convert]::ToBoolean($function_parameters[2])
    $checkbox_selected_msdefender.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # Microsoft Defender を更新する - ラベル作成
    [System.Windows.Forms.Label]$label_selected_msdefender = New-Object System.Windows.Forms.Label
    $label_selected_msdefender.Location = New-Object System.Drawing.Point(($form_size[0] *0.08),($form_size[1] * 0.175))
    $label_selected_msdefender.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_selected_msdefender.Text = 'Microsoft Defender を更新する'
    $label_selected_msdefender.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # アプリ を更新する - チェックボックスの作成
    [System.Windows.Forms.CheckBox]$checkbox_selected_application = New-Object System.Windows.Forms.CheckBox
    $checkbox_selected_application.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.28))
    $checkbox_selected_application.Size = New-Object System.Drawing.Size(20, 20)
    $checkbox_selected_application.Checked = [System.Convert]::ToBoolean($function_parameters[3])
    $checkbox_selected_application.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # アプリ を更新する - ラベル作成
    [System.Windows.Forms.Label]$label_selected_application = New-Object System.Windows.Forms.Label
    $label_selected_application.Location = New-Object System.Drawing.Point(($form_size[0] *0.08),($form_size[1] * 0.28))
    $label_selected_application.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_selected_application.Text = 'アプリ を更新する'
    $label_selected_application.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # アプリ を更新する - 確認ボタンの作成
    [System.Windows.Forms.Button]$btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Location = New-Object System.Drawing.Point(($form_size[0] * 0.820), ($form_size[1] * 0.28))
    # $btnCheck.Size = New-Object System.Drawing.Size(75,25)
    $btnCheck.Size = New-Object System.Drawing.Size(60,25)
    $btnCheck.Text = '確認'

    # Windows Update / アプリ 一括更新する - チェックボックスの作成
    [System.Windows.Forms.CheckBox]$checkbox_all_update = New-Object System.Windows.Forms.CheckBox
    $checkbox_all_update.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.385))
    $checkbox_all_update.Size = New-Object System.Drawing.Size(20, 20)
    $checkbox_all_update.Checked = [System.Convert]::ToBoolean($function_parameters[4])
    $checkbox_all_update.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # Windows Update / アプリ を一括更新する - ラベル作成
    [System.Windows.Forms.Label]$label_all_update = New-Object System.Windows.Forms.Label
    $label_all_update.Location = New-Object System.Drawing.Point(($form_size[0] *0.08),($form_size[1] * 0.385))
    $label_all_update.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_all_update.Text = '（オプション）Windows Update / アプリ を一括更新する'
    $label_all_update.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = '次へ'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($checkbox_selected_windowsupdate)
    $form.Controls.Add($label_selected_windowsupdate)
    $form.Controls.Add($checkbox_selected_msdefender)
    $form.Controls.Add($label_selected_msdefender)
    $form.Controls.Add($checkbox_selected_application)
    $form.Controls.Add($label_selected_application)
    $form.Controls.Add($checkbox_all_update)
    $form.Controls.Add($label_all_update)
    $form.Controls.Add($btnCheck)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # Windows Update を更新する ラベルの処理
    $label_selected_windowsupdate.add_click{
        $checkbox_selected_windowsupdate.Checked = !($checkbox_selected_windowsupdate.Checked)
    }

    # Microsoft Defender を更新する ラベルの処理
    $label_selected_msdefender.add_click{
        $checkbox_selected_msdefender.Checked = !($checkbox_selected_msdefender.Checked)
    }

    # アプリ を更新する ラベルの処理
    $label_selected_application.add_click{
        $checkbox_selected_application.Checked = !($checkbox_selected_application.Checked)
    }

    # アプリ を一括更新する ラベルの処理
    $label_all_update.add_click{
        $checkbox_all_update.Checked = !($checkbox_all_update.Checked)
    }

    # 確認ボタンの処理
    $btnCheck.add_click{
        # アクティブウィンドウをコマンドプロンプトとする
        [System.String]$cscode_filepath = "$($function_parameters[0])\source\csharp\ActiveWindow.cs"
        SwitchActiveWindow $cscode_filepath 'cmd'
        Write-Host 'winget のアップデート状況を確認中です...'
        Write-Host ''
        Write-Host ''

        # アップデート可能なアプリ確認
        (winget upgrade | Out-Host)
        Write-Host ''
        Write-Host ''
        Write-Host 'winget のアップデート状況を確認が完了しました。元の画面に戻り作業を再開してください。'
        Write-Host ''
        Write-Host ''
    }

    # フォーム表示
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.Int32]$max_retries = 3
    for ([System.Int32]$i=0; $i -le $max_retries; $i++) {
        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # 入力値のチェック
            [System.Boolean[]]$setting_parameters = @()
            $setting_parameters = @(
                [System.Convert]::ToBoolean($checkbox_selected_windowsupdate.Checked),
                [System.Convert]::ToBoolean($checkbox_selected_msdefender.Checked),
                [System.Convert]::ToBoolean($checkbox_selected_application.Checked),
                [System.Convert]::ToBoolean($checkbox_all_update.Checked)
            )
            $messagecode = ValidateInputValues $setting_parameters

            # チェック結果が正常の場合
            if ($messagecode -eq [MESSAGECODE]::Successful) {
                $form = $null
                break
            }
        }
        else {
            $setting_parameters = @()
            $form = $null
            break
        }
        # 再試行回数を超過前の処理
        if ($i -eq $max_retries) {
            $messagecode = [MESSAGECODE]::Error_MaxRetries
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '再試行回数の超過'
            ShowMessagebox $messagebox_messages $messagebox_title
            $setting_parameters = @()
            $form = $null
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function SettingInputValues: setting_parameters [${setting_parameters}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $setting_parameters
}

#################################################################################
# 処理名　 | RetrieveMessage
# 機能　　 | メッセージ内容を取得
#--------------------------------------------------------------------------------
# 戻り値　 | String（メッセージ内容）
# 引数　　 | target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function RetrieveMessage {
    Param (
        [MESSAGECODE]$target_code,
        [System.String]$append_message=''
    )

    [System.String]$return_messages = ''
    [System.String]$message = ''

    switch($target_code) {
        Successful                          {$message='正常終了';break}
        Abend                               {$message='異常終了';break}
        Cancel                              {$message='キャンセルしました。';break}
        Info_LoadedSettingfile              {$message='設定ファイルの読み込みが完了。';break}
        Info_SkipUpdate                     {$message='最新の状態である為、更新処理をスキップ。';break}
        Info_SkipMSDefenderInvalid          {$message='Microsoft Defenderが無効状態の為、更新処理をスキップ。';break}
        Info_SkipSelectWindowsUpdate        {$message='Windows Updateの更新でスキップが選択しました。';break}
        Info_SkipSelectMSDefender           {$message='Microsoft Defenderの更新でスキップが選択しました。';break}
        Info_SkipSelectWinget               {$message='winget の更新でスキップが選択しました。';break}
        Info_SkipExcludeWinget              {$message='更新対象外のアプリのため、更新をスキップします。';break}
        Confirm_ExecutionTool               {$message='ツールを実行します。';break}
        Confirm_ExecuteWindowsUpdate        {$message='Windows Updateの更新を実行します。';break}
        Confirm_ExecuteMSDedender           {$message='Microsoft Defenderの更新を実行します。';break}
        Confirm_ExecuteWinget               {$message='winget の更新を実行します。';break}
        Confirm_ExecuteWinget_Individual    {$message='個別に アプリ の更新を実行します。';break}
        Error_NotAdmin                      {$message='バッチファイルを“管理者として実行”してください。';break}
        Error_LoadingSettingfile            {$message='設定ファイルの読み込み処理でエラーが発生しました。';break}
        Error_EmptyTargetfolder             {$message='作業フォルダーが空で指定されています。';break}
        Error_NotCheckbox                   {$message='Windows Update / Microsoft Defender / アプリ いずれも未選択';break}
        Error_MaxRetries                    {$message='再試行回数を超過しました。';break}
        Error_InstallModules                {$message='モジュールのインストール／インポートの実行時にエラーが発生しました。';break}
        Error_GetWinUpdate                  {$message='Windows Updateの最新アップデートのチェックでエラーが発生しました。管理者として実行しているか確認してください。';break}
        Error_UpdateWinUpdate_all           {$message='Windows Updateの一括更新でエラーが発生しました。';break}
        Error_UpdateWinUpdate_Individual    {$message='Windows Updateの個別更新でエラーが発生しました。';break}
        Error_MSDefenderStatusCheck         {$message='Microsoft Defenderのステータスチェックでエラーが発生しました。';break}
        Error_UpdateMSDefender_Ex           {$message='Microsoft Defenderの更新処理で例外が発生しました。';break}
        Error_UpdateMSDefender_Returnerror  {$message='Microsoft Defenderの更新結果がエラーでした。';break}
        Error_UpdateWinget_Ex               {$message='wingetの更新処理（一括）で例外が発生しました。';break}
        Error_UpdateWinget_Individual_Ex    {$message='wingetの更新処理（個別）で例外が発生しました。';break}
        Error_UpdateWinget_Returnerror      {$message='wingetの更新結果がエラーでした。';break}
        Error_WingetNotInstall              {$message='winget 動作チェックでエラーが発生。アプリインストーラーを導入しているか確認してください。';break}
        Error_CheckWinget                   {$message='winget 動作チェックで例外エラーが発生しました。';break}
        Error_WingetUpgrade                 {$message='winget における更新アプリの確認でエラーが発生しました。';break}
        default                             {break}
    }

    $sbtemp=New-Object System.Text.StringBuilder
    @("${message}`r`n",`
      "${append_message}`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $return_messages = $sbtemp.ToString()

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RetrieveMessage: return_messages [${return_messages}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $return_messages
}

#################################################################################
# 処理名　 | ShowMessagebox
# 機能　　 | メッセージボックスの表示
#--------------------------------------------------------------------------------
# 戻り値　 | なし
# 引数　　 | target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function ShowMessagebox {
    Param (
        [System.String]$messages,
        [System.String]$title,
        [System.String]$level='Information'
        # 指定可能なレベル一覧（$level）
        #   None
        #   Hand
        #   Error
        #   Stop
        #   Question
        #   Exclamation
        #   Waring
        #   Asterisk
        #   Information
    )

    [System.Windows.Forms.DialogResult]$dialog_result = [System.Windows.Forms.MessageBox]::Show($messages, $title, "OK", $level)
    
    switch($dialog_result) {
        {$_ -eq [System.Windows.Forms.DialogResult]::OK} {
            break
        }
    }
}

#################################################################################
# 処理名　 | CreateExportFolder
# 機能　　 | 一時ファイルを格納するフォルダーを新規作成
#--------------------------------------------------------------------------------
# 戻り値　 | String（作成したフォルダー名。試行回数の超過もしくはエラーで作成できなかった場合は空文字を返す）
# 引数　　 | current_dir: 作業フォルダーのパス
# 　　　　 | foldername : 対象フォルダー名
# 　　　　 | max_retries: 最大のリトライ回数
#################################################################################
Function CreateExportFolder {
    Param (
        [System.String]$current_dir,
        [System.String]$foldername,
        [System.Int32]$max_retries=30
    )

    [System.String]$newfoldername = $foldername
    [System.Int32]$i = 0
    [System.String]$nowdate = (Get-Date).ToString("yyyyMMdd")
    [System.String]$number = ''
    for ($i=1; $i -le $max_retries; $i++) {
        # カウント数の数値を3桁で0埋めした文字列にする
        $number = "{0:000}" -f $i
        # 作成したいフォルダー名を生成
        $newfoldername = "${foldername}_${nowdate}-${number}"
        # 作成したいフォルダー名の存在チェック
        if (-Not (Test-Path "${current_dir}\${newfoldername}")) {
            break
        }

        # リトライ回数を超過し作成するフォルダー名を決定できなかった場合
        if ($i -eq $max_retries) {
            $newfoldername = ''
        }
    }

    [System.String]$newfolder_path = ''
    if ($newfoldername -ne '') {
        $newfolder_path = "${current_dir}\${newfoldername}"
        try {
            New-Item -Path "${newfolder_path}" -Type Directory > $null
        }
        catch {
            $newfolder_path = ''
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function CreateTempFolder: newfolder_path [${newfolder_path}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $newfolder_path
}

#################################################################################
# 処理名　 | ExtractByteSubstring
# 機能　　 | バイト数で文字列を抽出
#--------------------------------------------------------------------------------
# 戻り値　 | String（抽出した文字列）
# 引数　　 | target_str: 対象文字列
# 　　　　 | start     : 抽出開始するバイト位置
# 　　　　 | length    : 指定バイト数
#################################################################################
Function ExtractByteSubstring {
    Param (
        [System.String]$target_str,
        [System.Int32]$start,
        [System.Int32]$length
    )

    $encoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")

    # 文字列をバイト配列に変換
    [System.Byte[]]$all_bytes = $encoding.GetBytes($target_str)

    # 抽出するバイト配列を初期化
    $extracted_bytes = New-Object Byte[] $length

    # 指定されたバイト位置からバイト配列を抽出
    [System.Array]::Copy($all_bytes, $start, $extracted_bytes, 0, $length)

    # 抽出したバイトデータを文字列として返す
    return $encoding.GetString($extracted_bytes)
}

#################################################################################
# 処理名　 | InstallModules
# 機能　　 | リストのモジュールをインストール
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | install_modules : インストール対象のモジュール
#################################################################################
Function InstallModules {
    Param (
        [System.String[]]$function_parameters
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful

    [System.String[]]$modules_array = $function_parameters[0].Split(',')

    foreach ($module_name in $modules_array) {
        # インストールの有無をチェック
        if ($null -eq (Get-Module -Name "$module_name")) {
            # モジュールがない為、インストール
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-Module -Name "$module_name"
                Import-Module -Name "$module_name"

                Write-Host "$module_name has been installed." -ForegroundColor Cyan
                Write-Host ''
                Write-Host ''
            }
            catch {
                $messagecode = [MESSAGECODE]::Error_InstallModules
            }
        }
        else {
            # インストール済みの為、スキップ
            Write-Host "$module_name has been already installed."
            Write-Host ''
            Write-Host ''
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function InstallModules: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
#################################################################################
# 処理名　 | ExecuteWindowsUpdate
# 機能　　 | WindowsUpdateの実行
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | －
#################################################################################
Function ExecuteWindowsUpdate {
    Param (
        [System.Boolean]$all_update
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_messages = ''

    # 最新のアップデートをチェック
    try {
        # Get-WindowsUpdateコマンドレットのデータ型は、キートークンなどの情報がある為、データ型の指定は省略
        # System.Collections.ObjectModel.Collection`1[[System.Management.Automation.PSObject, System.Management.Automation, Version=7.3.10.500, Culture=neutral, PublicKeyToken=XXXXXXXXXXXXXXXX]]

        ### DEBUG ###
        if ($DEBUG_ON) {
            # デバッグだと管理者権限で実行できない為、アップデートがなかったものとして処理を進める
            $update_target = ''
        }
        else {
            $update_target = (Get-WindowsUpdate | Out-String)
            Write-Host $update_target
            Write-Host ''
            Write-Host ''
        }
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_GetWinUpdate
    }
    
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        if ($update_target -eq '') {
            $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Info_SkipUpdate)
            Write-Host $messagecode_messages
            Write-Host ''
            Write-Host ''

            return $messagecode
        }
    }

    # アップデート実行（アップデートがある場合）
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Confirm_ExecuteWindowsUpdate)
        # 対話式実行
        if (-Not($CONFIG_AUTO_MODE)) {
            if (ConfirmYesno $messagecode_messages) {
                if ($all_update) {
                    # 一括アップデート
                    try {
                        Install-WindowsUpdate -AcceptAll | Out-Host
                    }
                    catch {
                        $messagecode = [MESSAGECODE]::Error_UpdateWinUpdate_all
                    }
                }
                else {
                    # 個別アップデート
                    try {
                        Install-WindowsUpdate | Out-Host
                    }
                    catch {
                        $messagecode = [MESSAGECODE]::Error_WinUpdate_Individual
                    }
                }
            }
            else {
                $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Info_SkipSelectWindowsUpdate)
                Write-Host $messagecode_messages
                Write-Host ''
                Write-Host ''
            }
        }
        # 自動実行
        else {
            Write-Host $messagecode_messages
            Write-Host ''
            Write-Host ''
            if ($all_update) {
                # 一括アップデート
                try {
                    Install-WindowsUpdate -AcceptAll | Out-Host
                }
                catch {
                    $messagecode = [MESSAGECODE]::Error_UpdateWinUpdate_all
                }
            }
            else {
                # 個別アップデート
                try {
                    Install-WindowsUpdate | Out-Host
                }
                catch {
                    $messagecode = [MESSAGECODE]::Error_WinUpdate_Individual
                }
            }
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ExecuteWindowsUpdate: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
#################################################################################
# 処理名　 | ExecuteMicrosoftDefender
# 機能　　 | WindowsUpdateの実行
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | －
#################################################################################
Function ExecuteMicrosoftDefender {
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_messages = ''

    # Microsoft Defenderが有効状態かチェック
    try {
        [System.String]$status = (Get-Service -Name "WinDefend").Status
        [System.Boolean]$is_running = ($status -eq 'Running')
        if ($status -eq 'Running') {
            $is_running = $true
        }
        else {
            $messagecode = [MESSAGECODE]::Info_SkipMSDefenderInvalid
            $messagecode_messages = RetrieveMessage $messagecode
            Write-Host $messagecode_messages
            Write-Host ''
            Write-Host ''

            return [MESSAGECODE]::Successful
        }
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_MSDefenderStatusCheck
    }

    # アップデート実施するか有無
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Confirm_ExecuteMSDedender)
        # 対話式実行
        if (-Not($CONFIG_AUTO_MODE)) {
            if (-Not(ConfirmYesno $messagecode_messages)) {
                $messagecode = [MESSAGECODE]::Info_SkipSelectMSDefender
                $messagecode_messages = RetrieveMessage $messagecode
                Write-Host $messagecode_messages
                Write-Host ''
                Write-Host ''

                return [MESSAGECODE]::Successful
            }
        }
    }

    # アップデート実行と結果表示
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        # アップデート処理
        try {
            Update-MpSignature
        }
        catch {
            $messagecode = [MESSAGECODE]::Error_UpdateMSDefender_Ex
        }
        # 実行結果
        if (-Not($?)) {
            $messagecode = [MESSAGECODE]::Error_UpdateMSDefender_Returnerror
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ExecuteMicrosoftDefender: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　 | UpdateWingetIndividual
# 機能　　 | 個別にwinget upgrade
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | －
#################################################################################
Function UpdateWingetIndividual {
    Param (
        [System.String[]]$exclude_array
    )
    [System.String]$winget_result = (winget upgrade | Out-String)
    [System.String[]]$lines = $winget_result.Split([Environment]::NewLine)
    [System.String]$line = ''
    [System.String]$name =''
    [System.String]$id = ''
    [System.String]$version = ''
    [System.String]$available = ''
    [System.String]$source = ''
    
    [System.Int32]$header_index = $lines | Where-Object { $_.StartsWith("名前") } | ForEach-Object { $lines.IndexOf($_) }
    [System.Int32]$value_start = $header_index + 2
    # ID
    [System.String]$header_id_value = $lines[$header_index].Substring(0,$lines[$header_index].IndexOf("ID"))
    [System.Int32]$id_start = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($header_id_value)
    # Version
    [System.String]$header_version_value = $lines[$header_index].Substring(0,$lines[$header_index].IndexOf("バージョン"))
    [System.Int32]$version_start = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($header_version_value)
    # Available
    [System.String]$header_available_value = $lines[$header_index].Substring(0,$lines[$header_index].IndexOf("利用可能"))
    [System.Int32]$available_start = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($header_available_value)
    # Source
    [System.String]$header_source_value = $lines[$header_index].Substring(0,$lines[$header_index].IndexOf("ソース"))
    [System.Int32]$source_start = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($header_source_value)

    # [System.Int32]$unnecessary_char_digi = 48
    # [System.String[]]$columns = @()
    [System.Object[]]$winget_array = @()
    [System.Int32]$line_bytes = 0

    for ([System.Int32]$i = $value_start; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        $line_bytes = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($line)
        # 項目「ソース」の桁数以上ある場合に処理する
        if ($line_bytes -ge $source_start) {
            $name = (ExtractByteSubstring $line 0 $id_start).Trim()
            $id = (ExtractByteSubstring $line $id_start ($version_start - $id_start)).Trim()
            $version = (ExtractByteSubstring $line $version_start ($available_start - $version_start)).Trim()
            $available = (ExtractByteSubstring $line $available_start ($source_start - $available_start)).Trim()
            $source = (ExtractByteSubstring $line $source_start ($line_bytes - $source_start)).Trim()

            $winget_row = New-Object PSObject -Property @{
                Name = $name
                ID = $id
                Version = $version
                AvailableVersion = $available
                Source = $source
            }
            $winget_array += $winget_row
        }
    }

    # IDごとにアップデート
    [System.Boolean]$is_update = $true
    try {
        foreach ($update_item in $winget_array) {
            $is_update = $true
            # 値がない場合はスキップ
            if ($update_item.ID.Length -eq 0) {
                $is_update = $false
            }
            # 除外IDに該当する場合はスキップ
            else {
                foreach ($exclude_item in $exclude_array) {
                    if ($update_item.ID -eq $exclude_item) {
                        $is_update = $false
                        $sbtemp=New-Object System.Text.StringBuilder
                        @("`r`n",`
                          "　対象ID: [$($update_item.ID)]")|
                        ForEach-Object{[void]$sbtemp.Append($_)}
                        $append_message = $sbtemp.ToString()
                        $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Info_SkipExcludeWinget) $append_message
                        Write-Host $messagecode_messages
                        Write-Host ''
                        Write-Host ''
                        break
                    }
                }
            }
            if ($is_update) {
                $sbtemp=New-Object System.Text.StringBuilder
                @("`r`n",`
                  "　対象ID: [$($update_item.ID)]")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $append_message = $sbtemp.ToString()
                $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Confirm_ExecuteWinget_Individual) $append_message
                # 対話式実行
                if (-Not($CONFIG_AUTO_MODE)) {
                    if (ConfirmYesno $messagecode_messages) {
                        winget upgrade --id "$($update_item.ID)" | Out-Host
                        Write-Host ''
                        Write-Host ''

                        # 実行結果
                        if (-Not($?)) {
                            $messagecode = [MESSAGECODE]::Error_UpdateWinget_Returnerror
                            return $messagecode
                        }
                    }
                }
                # 自動実行
                Write-Host $messagecode_messages
                Write-Host ''
                Write-Host ''
                winget upgrade --id "$($update_item.ID)" | Out-Host
                Write-Host ''
                Write-Host ''

                # 実行結果
                if (-Not($?)) {
                    $messagecode = [MESSAGECODE]::Error_UpdateWinget_Returnerror
                    return $messagecode
                }
            }
        }
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_UpdateWinget_Individual_Ex
    }

    return $messagecode
}

#################################################################################
# 処理名　 | ExecuteWinget
# 機能　　 | wingetの実行
#--------------------------------------------------------------------------------
# 戻り値　 | MESSAGECODE（enum）
# 引数　　 | all_update: アプリ一括アップデートの有無（True: 一括アップデート、False: 個別アップデート）
# 注意事項 | Appxモジュールではアップデートするコマンドがなかったため、wingetを使用。
# 　　　　 | なお、winget upgrade でアップデートできるのは、ソースが「winget」が対象となる。
# 　　　　 | ソースが「msstore」や「（空欄）」のアプリは更新できない。
# 機能拡張 | winget pin add で更新対象外を指定できるので対象外リストの機能を追加した方が便利かも。
#################################################################################
Function ExecuteWinget {
    Param (
        [System.String[]]$exclude_id,
        [System.Boolean]$all_update
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_messages = ''

    # wingetコマンドを実行しアプリインストーラーの導入状態を確認
    try {
        [void](winget)
        if ($? -eq $false) {
            $messagecode = [MESSAGECODE]::Error_WingetNotInstall
        }
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_CheckWinget
    }

    # アップデート有無の確認
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        try {
            (winget upgrade | Out-Host)
            Write-Host ''
            Write-Host ''
        }
        catch {
            $messagecode = [MESSAGECODE]::Error_WingetUpgrade
        }
    }

    # アップデートを実施するか
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Confirm_ExecuteWinget)
        # 対話式実行
        if (-Not($CONFIG_AUTO_MODE)) {
            if (-Not(ConfirmYesno $messagecode_messages)) {
                $messagecode = [MESSAGECODE]::Info_SkipSelectWinget
                $messagecode_messages = RetrieveMessage $messagecode
                Write-Host $messagecode_messages
                Write-Host ''
                Write-Host ''

                return [MESSAGECODE]::Successful
            }
        }
    }
    
    # アップデート実行
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        if ($all_update) {
            try {
                # 一括アップデート
                winget upgrade --all | Out-Host
                Write-Host ''
                Write-Host ''
            }
            catch {
                $messagecode = [MESSAGECODE]::Error_UpdateWinget_Ex
            }
            # 実行結果
            if (-Not($?)) {
                $messagecode = [MESSAGECODE]::Error_UpdateWinget_Returnerror
            }
        }
        else {
            # 個別アップデート
            $messagecode = UpdateWingetIndividual $exclude_id
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ExecuteWinget: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
### Function <--- 終了 ---

### Main process --- 開始 --->
#################################################################################
# 処理名　 | メイン処理
# 機能　　 | 同上
#--------------------------------------------------------------------------------
# 　　　　 | -
#################################################################################
# 初期設定
#   メッセージ関連
[MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
[System.String]$messagecode_messages = ''
[System.String]$append_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 管理者権限であることを確認
if (-Not(isAdminPowerShell)) {
    $messagecode = [MESSAGECODE]::Error_NotAdmin

    # デバッグだと管理者権限で実行できない為、管理者権限で起動したものとして処理を進める
    if ($DEBUG_ON) {
        $messagecode = [MESSAGECODE]::Successful
    }
}

# 文字コードの変更
# コンソール出力する文字コードを「UTF-8」に変更
if ($messagecode -eq [MESSAGECODE]::Successful) {
    SetPsOutputEncoding 'utf8'
    Write-Host ''
    Write-Host ''
}

# 設定ファイルの読み込み
# ディレクトリの取得
if ($messagecode -eq [MESSAGECODE]::Successful) {
    [System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
    Set-Location $current_dir'\..\..'
    [System.String]$root_dir = (Convert-Path .)

    # Configファイルのフルパスを作成  
    $sbtemp=New-Object System.Text.StringBuilder
    @("${current_dir}",`
        '\',`
        "${c_config_file}")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    [System.String]$config_fullpath = $sbtemp.ToString()

    # 読み込み処理
    try {
        [System.Collections.Hashtable]$config = (Get-Content $config_fullpath -Raw -Encoding UTF8).Replace('\','\\') | ConvertFrom-StringData

        # 変数に格納
        [System.String]$CONFIG_INSTALL_MODULES=(RemoveDoubleQuotes($config.install_modules))
        [System.Boolean]$CONFIG_SELECETED_WINDOWSUPDATE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.selected_windowsupdate)))
        [System.Boolean]$CONFIG_SELECETED_MSDEFENDER=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.selected_msdefender)))
        [System.Boolean]$CONFIG_SELECETED_APPLICATION=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.selected_application)))
        [System.Boolean]$CONFIG_ALL_UPDATE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.all_update)))
        [System.String[]]$CONFIG_EXCLUDE_ID=(RemoveDoubleQuotes($config.exclude_id)).Split(',')
        [System.Boolean]$CONFIG_AUTO_MODE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.auto_mode)))

        # 通知
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
            "対象ファイル: [${config_fullpath}]")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Info_LoadedSettingfile) $append_message
        Write-Host $messagecode_messages
        Write-Host ''
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_LoadingSettingfile
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
            "エラーの詳細: [${config_fullpath}$($_.Exception.Message)]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $messagecode_messages = RetrieveMessage $messagecode $append_message
    }
}

# 入力値の設定・検証
if ($messagecode -eq [MESSAGECODE]::Successful) {
    [System.Object[]]$function_parameters = @()
    [System.Boolean[]]$setting_parameters = @()
    # 対話式実行
    if (-Not($CONFIG_AUTO_MODE)) {
        # 対話式の場合：入力値を画面で設定（入力値の検証を含む）
        $function_parameters = @(
            $root_dir,
            $CONFIG_SELECETED_WINDOWSUPDATE,        # Windows Update 更新の有無
            $CONFIG_SELECETED_MSDEFENDER,           # Microsoft Defender 更新の有無
            $CONFIG_SELECETED_APPLICATION,          # winget upgrade 更新の有無
            $CONFIG_ALL_UPDATE                      # 一括更新の有無
        )
        
        $setting_parameters = SettingInputValues $function_parameters
        if ($null -eq $setting_parameters) {
            $messagecode = [MESSAGECODE]::Cancel
        }
    }
    # 自動実行
    else {
        # 自動実行の場合：入力値の検証のみ
        #   入力値のチェック
        [System.Boolean[]]$setting_parameters = @()
        $setting_parameters = @(
            $CONFIG_SELECETED_WINDOWSUPDATE,        # Windows Update 更新の有無
            $CONFIG_SELECETED_MSDEFENDER,           # Microsoft Defender 更新の有無
            $CONFIG_SELECETED_APPLICATION,          # winget upgrade 更新の有無
            $CONFIG_ALL_UPDATE                      # 一括更新の有無
        )
        $messagecode = ValidateInputValues $setting_parameters
    }
}

# 実行有無の確認
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Confirm_ExecutionTool)
    # 対話式実行
    if (-Not($CONFIG_AUTO_MODE)) {
        if (-Not(ConfirmYesno $messagecode_messages)) {
            $messagecode = [MESSAGECODE]::Cancel
        }
    }
}

# モジュールのインストール
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $function_parameters = @(
        $CONFIG_INSTALL_MODULES     # インストールするモジュール
    )

    # モジュールのインストール（未導入の場合）
    $messagecode = InstallModules $function_parameters
}

# Windows Update
if ($messagecode -eq [MESSAGECODE]::Successful) {
    if ($setting_parameters[0]) {
        $messagecode = ExecuteWindowsUpdate $setting_parameters[3]
    }
}

# Microsoft Defender
if ($messagecode -eq [MESSAGECODE]::Successful) {
    if ($setting_parameters[1]) {
        $messagecode = ExecuteMicrosoftDefender
    }
}

# インストーラー版アプリ（source = winget）
if ($messagecode -eq [MESSAGECODE]::Successful) {
    if ($setting_parameters[2]) {
        $messagecode = ExecuteWinget $CONFIG_EXCLUDE_ID $setting_parameters[3]
    }
}

# コンソール出力する文字コードを「初期値に戻す」に変更
if ($messagecode -eq [MESSAGECODE]::Successful) {
    SetPsOutputEncoding 'reset_encoding'
    Write-Host ''
    Write-Host ''
}

#   処理結果の表示
[System.String]$append_message = ''
$sbtemp=New-Object System.Text.StringBuilder
if ($messagecode -eq [MESSAGECODE]::Successful) {
    @("`r`n",`
      "メッセージコード: [${messagecode}]")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_messages = RetrieveMessage $messagecode $append_message
    Write-Host $messagecode_messages
    Write-Host ''
    Write-Host ''
}
else {
    @("`r`n",`
      "メッセージコード: [${messagecode}]",`
      $messagecode_messages)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Abend) $append_message
    Write-Host $messagecode_messages -ForegroundColor Red
    Write-Host ''
    Write-Host ''
}
# 終了

# PowerShellスクリプト上で一時停止する必要がある場合
# [System.Windows.Forms.MessageBox]::Show("OK Button to Continue.")

exit $messagecode
### Main process <--- 終了 ---
