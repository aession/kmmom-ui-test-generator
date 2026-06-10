---
name: kmmom-ui-test-generator
description: Use when Codex needs to operate or inspect the KMMOM system in local Google Chrome, traverse a left-menu module, identify and safely click visible UI buttons/filters/dropdowns/dialog entries/table controls, generate a page button/control click test Excel record, or convert that record into MeterSphere-format test cases.
---

# KMMOM UI Test Generator

## Use This Skill When

Use this skill for requests such as:

- 登录并操作 KMMOM 系统页面。
- 遍历某个左侧菜单模块，例如“制造执行”“质量管理”“设备管理”。
- 识别页面按钮、筛选条件、下拉框、日期控件、弹窗/抽屉入口、表格控件。
- 生成“界面按钮/控件点击测试表”。
- 将页面控件点击记录转换为 MeterSphere 导入格式测试用例。
- 把一次性页面巡检工作沉淀为可重复执行的 UI 自动化辅助脚本。

## Standard Workflow

1. Confirm the user requested target module or route scope. Default module is `制造执行`.
2. Confirm Chrome is available. If Chrome is not running, ask the user before launching it.
3. Connect to local Chrome through CDP, usually `http://127.0.0.1:9223/json/list`.
4. Open KMMOM, log in or reuse an existing login state.
5. Read KMMOM menu data from `localStorage`, then select leaf pages under the requested module.
6. Visit each page and identify visible UI areas:
   - 筛选区
   - 页面条件区
   - 工具栏
   - 表格区
   - 弹窗/抽屉
   - 视图区
   - 公共浮动区
7. Safely exercise controls: expand filters, open dropdowns, open/close dialogs, input/clear non-saving test text, click non-risk buttons.
8. During every click or input interaction, observe and record abnormal responses and visible prompts:
   - toast/message/notification text
   - alert banners
   - modal/dialog/drawer title, body text, and footer button text
   - permission/login-expired/interceptor prompts
   - API failure, loading failure, blank page, error page, or empty-data prompts
   - whether the prompt blocks subsequent operation and whether it can be closed safely
9. Do not execute final business-changing actions in production.
10. Save execution records under `outputs/04_测试执行/02_执行记录/`.
11. If requested, convert the execution record into MeterSphere import cases under `outputs/03_测试用例/03_MeterSphere导入/`.

## Safety Boundary

Production-like KMMOM environments must default to `safe_mode=true`.

Never execute final actions containing these texts unless the user explicitly confirms a safe test environment and data cleanup plan:

- 保存
- 提交
- 确定 / 确 定
- 删除
- 导出
- 下载
- 初始化
- 发布
- 下发
- 通过
- 驳回
- 发送消息

For risky controls, record visibility, enabled/disabled state, expected result, dialog content, and risk reason. Do not click the final action.

Read [references/safety-boundary.md](references/safety-boundary.md) when deciding whether an action is safe.

## Scripts

Prefer the bundled scripts instead of rewriting automation code.

### Launch Chrome

Use only after the user has allowed opening a controllable Chrome window.

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\launch_chrome.ps1" `
  -Url "http://192.168.30.69:40000" `
  -RemoteDebuggingPort 9223 `
  -UserDataDir "D:\Personal\Documents\KMMOM系统管理后台\chrome_kmmom_profile"
```

### Generate UI Click Test Record

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" `
  --base-url "http://192.168.30.69:40000" `
  --module "制造执行" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台" `
  --debug-url "http://127.0.0.1:9223/json/list" `
  --safe-mode
```

The script outputs:

```text
outputs/04_测试执行/02_执行记录/KMMOM-<模块名>模块-界面按钮点击测试表-YYYYMMDD.xlsx
```

### Generate MeterSphere Cases

Default granularity is page-level: one interface/page becomes one MeterSphere case, and the page's buttons/controls become ordered case steps. Use `--granularity control` only when the user explicitly wants one control per case.

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\generate_metersphere_cases.py" `
  --source "D:\Personal\Documents\KMMOM系统管理后台\outputs\04_测试执行\02_执行记录\KMMOM-制造执行模块-界面按钮点击测试表-YYYYMMDD.xlsx" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台"
```

The script outputs:

```text
outputs/03_测试用例/03_MeterSphere导入/KMMOM-<模块名>模块-页面按钮控件点击测试用例-MeterSphere导入-YYYYMMDD.xlsx
```

For fine-grained control-level cases:

```powershell
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\generate_metersphere_cases.py" `
  --source "<界面按钮点击测试表.xlsx>" `
  --granularity control
```

## Output Standards

Execution record Excel must include:

- `按钮点击测试表`
- `页面汇总`
- `测试说明`

Main record fields must include:

- 序号
- 模块
- 菜单路径
- 所属页面
- 路由地址
- 区域
- 按钮/控件
- 控件类型
- 点击前置
- 预期结果
- 实测状态
- 展开/弹窗/选项内容
- 异常/弹窗提示
- 是否执行点击
- 风险说明

MeterSphere case Excel must use these columns exactly:

- 用例名称
- 所属模块
- 标签
- 前置条件
- 步骤描述
- 预期结果
- 编辑模式
- 备注
- 用例状态
- 责任人
- 用例等级

Default case design rule:

- One page/interface = one MeterSphere test case.
- Each visible button/control/filter/dropdown/dialog entry/table control on that page = one or more test steps.
- Do not create one case per control unless the user explicitly asks for a fine-grained traceability version.

Read [references/output-rules.md](references/output-rules.md) and [references/metersphere-case-format.md](references/metersphere-case-format.md) when generating or validating output files.

## Common Commands

List pages for a module without clicking:

```powershell
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" --module "制造执行" --list-pages-only
```

Traverse another module:

```powershell
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" --module "质量管理" --safe-mode
```

## Notes

- If command-line Chinese text appears garbled in PowerShell, set `$env:PYTHONIOENCODING='utf-8'`.
- Prefer parameterized module names over hardcoded routes.
- For broad system runs, test one module at a time. This produces smaller, reviewable artifacts and avoids unnecessary production load.
