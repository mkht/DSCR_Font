﻿#
# モジュール 'DSCR_Font' のモジュール マニフェスト
#
# 生成者: mkht
#
# 生成日: 2016/11/27
#

@{

# このマニフェストに関連付けられているスクリプト モジュール ファイルまたはバイナリ モジュール ファイル。
# RootModule = ''

# このモジュールのバージョン番号です。
ModuleVersion = '0.6.0'

# このモジュールを一意に識別するために使用される ID
GUID = 'c8a4a363-3b1f-444f-913b-3362f69e6c27'

# このモジュールの作成者
Author = 'mkht'

# このモジュールの会社またはベンダー
CompanyName = ''

# このモジュールの著作権情報
Copyright = '(c) 2017 mkht. All rights reserved.'

# このモジュールの機能の説明
Description = 'PowerShell DSC Resource to add / remove Font.'

# このモジュールに必要な Windows PowerShell エンジンの最小バージョン
PowerShellVersion = '5.0'

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# RootModule/ModuleToProcess に指定されているモジュールの入れ子になったモジュールとしてインポートするモジュール
# NestedModules = @()

# このモジュールからエクスポートする関数です。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートする関数がない場合は、エントリを削除しないで空の配列を使用してください。
FunctionsToExport = @()

# このモジュールからエクスポートするコマンドレットです。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートするコマンドレットがない場合は、エントリを削除しないで空の配列を使用してください。
CmdletsToExport = @()

# このモジュールからエクスポートする変数
VariablesToExport = '*'

# このモジュールからエクスポートするエイリアスです。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートするエイリアスがない場合は、エントリを削除しないで空の配列を使用してください。
AliasesToExport = @()

# このモジュールからエクスポートする DSC リソース
DscResourcesToExport = @('cFont')

# このモジュールからエクスポートされたコマンドの既定のプレフィックス。既定のプレフィックスをオーバーライドする場合は、Import-Module -Prefix を使用します。
# DefaultCommandPrefix = ''

# RootModule/ModuleToProcess に指定されているモジュールに渡すプライベート データ。これには、PowerShell で使用される追加のモジュール メタデータを含む PSData ハッシュテーブルが含まれる場合もあります。
PrivateData = @{

    PSData = @{

        # このモジュールに適用されているタグ。オンライン ギャラリーでモジュールを検出する際に役立ちます。
        Tags = ('DesiredStateConfiguration','DSC','DSCResource', 'Font')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/mkht/DSCR_Font/blob/master/LICENSE'

        # このプロジェクトのメイン Web サイトの URL。
        ProjectUri = 'https://github.com/mkht/DSCR_Font'

        # このモジュールを表すアイコンの URL。
        # IconUri = ''

        # このモジュールの ReleaseNotes
        # ReleaseNotes = ''

    } # PSData ハッシュテーブル終了

} # PrivateData ハッシュテーブル終了

}

