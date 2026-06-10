# KMMOM UI 测试技能仓库

本仓库用于分享和复现 `kmmom-ui-test-generator` 技能。

该技能用于 KMMOM/MOM 系统 UI 页面遍历、控件采集、按钮点击记录生成和异常线索收集。它适合测试人员在回归测试、版本升级验证、页面控件盘点时使用。

重要说明：

> 这个技能是 UI 测试辅助工具，不是自动缺陷判定工具。  
> 控件清单和点击记录可作为测试覆盖资料；异常汇总必须人工复核后，才能决定是否提缺陷。

## 1. 仓库内容

```text
kmmom-ui-test-generator/
├─ README.md
├─ requirements.txt
├─ .gitignore
├─ scripts/
│  └─ install-skill.ps1
├─ docs/
│  ├─ usage-guide.md
│  └─ github-upload.md
└─ skills/
   └─ kmmom-ui-test-generator/
      ├─ SKILL.md
      ├─ agents/
      ├─ assets/
      ├─ references/
      └─ scripts/
```

主要文件说明：

| 文件/目录 | 作用 |
| --- | --- |
| `skills/kmmom-ui-test-generator/SKILL.md` | 技能主说明文件 |
| `skills/kmmom-ui-test-generator/scripts/` | UI 遍历、记录生成、MeterSphere 转换脚本 |
| `skills/kmmom-ui-test-generator/references/` | 输出规则、安全边界、MeterSphere 格式说明 |
| `skills/kmmom-ui-test-generator/assets/config.example.json` | 示例配置文件 |
| `docs/usage-guide.md` | 测试人员使用说明 |
| `docs/github-upload.md` | GitHub 上传说明 |
| `scripts/install-skill.ps1` | 本地安装脚本 |

## 2. 使用前提

使用前请确认：

- 当前电脑可以访问 KMMOM/MOM 系统；
- 当前账号具备目标模块权限；
- 已安装 Python；
- 已安装 Git；
- 已安装 Chrome；
- 已登录 KMMOM/MOM 系统，或具备可登录账号；
- 默认使用安全模式，不执行保存、提交、删除、导出、审核等最终业务动作。

## 3. 安装 Python 依赖

在仓库根目录执行：

```powershell
pip install -r requirements.txt
```

当前依赖：

```text
openpyxl>=3.1.0
```

## 4. 安装 Skill 到本机 Codex

推荐使用安装脚本：

```powershell
cd "D:\Personal\Documents\KMMOM系统管理后台\github_package\kmmom-ui-test-generator"
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

该脚本会把 skill 复制到：

```text
C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator
```

安装完成后，建议重启 Codex。

## 5. 配置登录密码

仓库中不会保存真实密码。

如需让脚本使用环境变量读取密码，可执行：

```powershell
cd "D:\Personal\Documents\KMMOM系统管理后台\github_package\kmmom-ui-test-generator"
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1 -SetPassword
```

脚本会提示输入 KMMOM 密码，并保存到当前 Windows 用户的环境变量：

```text
KMMOM_PASSWORD
```

密码不会写入 Git 仓库。

## 6. 当前默认环境配置

当前包内已按你的环境写入默认配置：

| 配置项 | 当前值 |
| --- | --- |
| KMMOM 地址 | `http://192.168.30.69:40000` |
| Chrome 调试地址 | `http://127.0.0.1:9223/json/list` |
| Chrome 调试端口 | `9223` |
| 工作区 | `D:\Personal\Documents\KMMOM系统管理后台` |
| Skill 安装目录 | `C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator` |
| 默认模块 | `制造执行` |
| 安全模式 | `true` |

## 7. 启动可控 Chrome

执行：

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\launch_chrome.ps1" `
  -Url "http://192.168.30.69:40000" `
  -RemoteDebuggingPort 9223 `
  -UserDataDir "D:\Personal\Documents\KMMOM系统管理后台\chrome_kmmom_profile"
```

启动后，在该 Chrome 中登录 KMMOM/MOM 系统。

## 8. 遍历单个模块

示例：遍历制造执行模块。

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" `
  --base-url "http://192.168.30.69:40000" `
  --module "制造执行" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台" `
  --debug-url "http://127.0.0.1:9223/json/list" `
  --safe-mode
```

输出目录：

```text
D:\Personal\Documents\KMMOM系统管理后台\outputs\04_测试执行\02_执行记录
```

## 9. 只列出模块页面

如果只想确认模块下有哪些页面，不执行点击：

```powershell
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" `
  --module "制造执行" `
  --list-pages-only
```

## 10. 转换 MeterSphere 用例

如果需要把点击记录转换成 MeterSphere 导入格式：

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\generate_metersphere_cases.py" `
  --source "D:\Personal\Documents\KMMOM系统管理后台\outputs\04_测试执行\02_执行记录\KMMOM-制造执行模块-界面按钮点击测试表-YYYYMMDD.xlsx" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台"
```

默认是一页一个测试用例，页面内控件作为步骤。

## 11. 安全边界

默认不执行以下最终业务动作：

- 保存；
- 提交；
- 删除；
- 导出；
- 下载；
- 发布；
- 下发；
- 审核；
- 通过；
- 驳回；
- 初始化；
- 发送消息。

对高风险控件，只记录可见性、启用状态、预期结果、弹窗内容和风险说明，不点击最终确认。

## 12. 异常结果复核原则

异常汇总只是线索，不是缺陷清单。

需要重点复核：

- 登录失效；
- 网络连接重置；
- 服务异常；
- 接口失败；
- 页面空白；
- 权限提示；
- 业务规则提示是否触发合理。

不要直接提缺陷的情况：

- 正常删除确认弹窗；
- 正常取消/确定弹窗；
- 弹窗残留；
- 页面内容残留；
- 坐标元素未找到；
- `Runtime.evaluate` 等自动化执行错误；
- 正常页面跳转。

复核原则：

> 能人工稳定复现的问题，才进入缺陷判断。  
> 只在自动化记录中出现、人工无法复现的问题，先标记为待复测。

## 13. 上传 GitHub

上传说明见：

[docs/github-upload.md](docs/github-upload.md)

建议仓库设置为 private。

## 14. 给测试人员的标准使用指令

```text
基于 KMMOM UI 测试 skill，遍历【模块名称】模块下全部页面，
采集页面内所有按钮、筛选控件、下拉控件、输入框、日期控件、表格控件、页签和弹窗入口，
输出界面按钮点击测试记录表、完整控件清单和异常汇总。
开启安全模式，不执行保存、提交、删除、导出、下载、审核、发布等最终业务动作。
异常结果仅作为复核线索，不直接作为缺陷结论。
无需生成测试用例。
```

## 15. 总结

这个技能的最佳使用方式是：

> 自动化负责采集和记录，人工负责复核和判断。

它可以提高 UI 控件盘点和回归前检查效率，但不能替代人工测试和缺陷判断。
