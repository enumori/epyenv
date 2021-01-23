# epyenv
Simple Python Windows x86-64 embeddable installer

# Installation
powershellを起動して以下のコマンドを入力します。
```
Set-ExecutionPolicy RemoteSigned -scope Process
Invoke-WebRequest -Uri "https://github.com/enumori/epyenv/releases/download/2021.01.23/epyenv.zip" -OutFile .\epyenv.zip
Expand-Archive -Path .\epyenv.zip -DestinationPath $env:USERPROFILE
Remove-Item .\epyenv.zip
Rename-Item  $env:USERPROFILE\ndenv  $env:USERPROFILE\.ndenv
$path = [Environment]::GetEnvironmentVariable("PATH", "User")
$path = "$env:USERPROFILE\.epyenv;" + $path
[Environment]::SetEnvironmentVariable("PATH", $path, "User")
```
powershellやコマンドプロンプトを起動するとepyenvが使用できます。

# Command Reference
| 実行内容 | コマンド|
| --- | --- |
| インストール可能なPython Windows embeddable packageバージョンのリスト | epyenv list |
| カレントディレクトリへのインストール | epyenv install バージョン |
| 指定したディレクトリへのインストール| epyenv install バージョン -out ディレクトリパス |

# 使い方
## カレントディレクトリにPython Windows x86-64 embeddableをインストールする
```
PS > epyenv install 3.6.0-amd64
```
## C:\python_programsにPython Windows x86-64 embeddableをインストールする

```
PS > epyenv install 3.6.0-amd64 -o C:\python_programs
```
