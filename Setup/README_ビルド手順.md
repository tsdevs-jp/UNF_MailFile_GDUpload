# UNF MailFile GD Upload — インストーラービルド手順

## インストーラーのビルド

### 1. 前提ツール

| ツール | バージョン | 入手先 |
|--------|-----------|--------|
| Visual Studio 2022/2026 | 任意 | — |
| Inno Setup 6 | 6.x 系 | https://jrsoftware.org/isdl.php |

---

### 2. ビルド手順

```
1. Visual Studio で「Release」構成のビルドを実行する
   → bin\Release\ にアドインファイル一式が生成される

2. Inno Setup Compiler を起動し、Setup\setup.iss を開く
   （または ISCC.exe でコマンドラインビルドも可）
   > "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" Setup\setup.iss

3. Setup\output\UNF_MailFile_GDUpload_Setup_1.0.0.exe が生成される
```

---

### 3. インストーラーの動作

| 項目 | 内容 |
|------|------|
| インストール先 | `C:\Program Files\Unfake\UNF_MailFile_GDUpload\` |
| Unfake フォルダ | 存在しない場合は自動作成 |
| 必要権限 | 管理者権限 |
| アドイン登録 | HKCU / HKLM レジストリへ自動登録 |
| 署名証明書 | Trusted Publishers / Root CA ストアへ自動インポート |
| 前提条件チェック | .NET Framework 4.8・VSTO 4.0 Runtime がない場合はインストール中断 |

---

### 4. 配布 PC での前提条件

配布先 PC に以下がインストールされていること。  
インストーラーが起動前にチェックし、不足している場合はエラーメッセージで URL を案内します。

| コンポーネント | 入手先 |
|---------------|--------|
| .NET Framework 4.8 | https://dotnet.microsoft.com/download/dotnet-framework/net48 |
| Visual Studio 2010 Tools for Office Runtime | https://www.microsoft.com/download/details.aspx?id=56961 |

---

### 5. アンインストール

「コントロールパネル → プログラムのアンインストール」から **UNF MailFile GD Upload** を選択。  
- アドインのレジストリ登録が削除されます。  
- 署名証明書がストアから削除されます。  
- `Unfake` フォルダが空になった場合は自動削除されます。

---

### 6. バージョン更新時

1. `AssemblyInfo.cs` の `AssemblyVersion` / `AssemblyFileVersion` を更新  
2. `setup.iss` 冒頭の `#define AppVersion` を同じ番号に更新  
3. Visual Studio で Release ビルド実行  
4. Inno Setup でインストーラー再ビルド  
