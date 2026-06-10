# GitHub 上传说明

本文说明如何把当前 `kmmom-ui-test-generator` 技能包上传到 GitHub。

当前本地仓库包路径：

```text
D:\Personal\Documents\KMMOM系统管理后台\github_package\kmmom-ui-test-generator
```

当前 GitHub 用户：

```text
aession
```

推荐目标仓库：

```text
https://github.com/aession/kmmom-ui-test-generator
```

## 1. 上传前检查

上传前确认仓库包中不包含：

- 浏览器 profile；
- 执行输出 Excel；
- 异常截图；
- 日志文件；
- zip 压缩包；
- `__pycache__`；
- `.pyc` 文件；
- 明文密码；
- 生产数据。

当前包中保留了你的用户目录和内网地址，方便拿到后直接按本机环境使用：

```text
C:\Users\Administrator
D:\Personal\Documents\KMMOM系统管理后台
http://192.168.30.69:40000
```

密码不写入仓库。需要在本机执行安装脚本时输入，并保存到 `KMMOM_PASSWORD` 环境变量。

## 2. 在 GitHub 创建仓库

打开：

```text
https://github.com/new
```

填写：

```text
Repository name: kmmom-ui-test-generator
Visibility: Private
```

建议选择 private。

不要勾选：

```text
Add a README file
Add .gitignore
Choose a license
```

因为这些文件已经在本地包里准备好了。

点击：

```text
Create repository
```

## 3. 使用 Git 命令上传

进入本地包目录：

```powershell
cd "D:\Personal\Documents\KMMOM系统管理后台\github_package\kmmom-ui-test-generator"
```

初始化仓库：

```powershell
git init
git add .
git commit -m "Add KMMOM UI test generator skill"
git branch -M main
git remote add origin https://github.com/aession/kmmom-ui-test-generator.git
git push -u origin main
```

## 4. 如果远程仓库已经存在内容

如果 GitHub 上的仓库已经有文件，推荐使用临时目录方式上传，避免覆盖错。

```powershell
cd "D:\Personal\Documents\KMMOM系统管理后台\github_package"

git clone https://github.com/aession/kmmom-ui-test-generator.git github-upload-work

Copy-Item -Recurse -Force `
  ".\kmmom-ui-test-generator\*" `
  ".\github-upload-work\"

cd ".\github-upload-work"

git add .
git commit -m "Update KMMOM UI test generator skill package"
git push
```

## 5. 如果命令行没有 GitHub 权限

如果执行 `git push` 时提示需要登录，按提示完成 GitHub 登录。

如果命令行无法连接 GitHub，可以改用 GitHub Desktop 或网页上传。

网页上传时，不要上传 ZIP 文件本身，而是上传解压后的目录内容。

## 6. 上传成功后别人如何使用

别人 clone 仓库：

```powershell
git clone https://github.com/aession/kmmom-ui-test-generator.git
cd kmmom-ui-test-generator
```

安装依赖：

```powershell
pip install -r requirements.txt
```

安装 skill：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

安装 skill 并配置本机密码环境变量：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1 -SetPassword
```

安装完成后重启 Codex。

## 7. 注意事项

- 仓库建议设置为 private；
- 不要把真实密码提交到 GitHub；
- 不要上传执行结果和截图；
- 不要上传浏览器 profile；
- 异常汇总只能作为复核线索；
- 生产或类生产环境必须使用安全模式。
