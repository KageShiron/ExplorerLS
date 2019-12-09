# ExplorerLS
現在開かれているエクスプローラのパスを表示します。

## 引数
<dl>
    <dt>expls.exe last</dl>
    <dd>最後にアクティブになったエクスプローラのパスを出力します</dd>
    <dt>expls.exe all</dl>
    <dd>すべてのエクスプローラのパスの一覧を表示します。</dd>
    <dt>expls.exe</dt>
    <dd>expls.exe lastと同様</dd>
    <dt>expls.exe help</dt>
    <dd>ヘルプ</dd>
</dl>

## その他仕様
* パスを持たないエクスプローラは列挙されません(仮想フォルダ)
* iexplore.exeのパスは列挙されません

# Usage
## Basic Usage
```
PS> expls.exe
C:\s

PS> expls.exe all
C:\s
C:\Users\minaduki\Desktop
D:\gd\pics
```

## Advanced Usage
Profile.ps1に関数を定義しておくのもおすすめです。

### PowerShellで最後にアクティブとなったウィンドウに移動
```
PS> explorels.exe | Set-Location
```

### エクスプローラのパス一覧から選択して移動
pecoを使用 [peco/peco: Simplistic interactive filtering tool](https://github.com/peco/peco)

```
PS> expls.exe all | peco | Set-Location
```

# Build
```
> msbuild /p:Configuration=Release
> msbuild /t:ILMerge
```

# おまけ
## 以前使っていたPowerShell
```powershell
function cde() {
    $shell = New-Object -ComObject Shell.Application
    $exps = $shell.Windows() | Select-Object @{Name = "Name"; expression = {$_.LocationName} } , @{Name = "Path"; expression = {([uri]$_.LocationURL).LocalPath} } | 
        Where-Object {$null -ne $_.Path}
    $exps | ForEach-Object { "({0}) {1}`0{1}" -f $_.Name, $_.Path} | peco --null | Set-Location
}
```

## 既存の問題
* `EnumWindows`が列挙する順番は本来保証されていませんが、アクティブになった順に列挙されるという挙動を期待しています。lastが本当に最後にアクティブだったウィンドウを返さない可能性がありますが、現状では仕様です。
* .NET Coreではなぜか`Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ShellWindows))`が成功しないため、.NET Frameworkを使っています。OleInitializeとか、スレッドアパートメントの変更とかはやってみたのですが、うまくいきません。.NET Coreで動いた方はお知らせください。


# License
[MIT License](./LICENSE.txt)
