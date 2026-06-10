# KMMOM UI Test Generator Skill

KMMOM UI Test Generator is a Codex skill for traversing KMMOM/MOM web pages, collecting visible UI controls, safely exercising low-risk interactions, and generating Excel execution records.

It is intended for KMMOM/MOM testing teams that need to:

- collect buttons, filters, dropdowns, inputs, date pickers, table controls, tabs, dialogs, and drawers;
- generate UI click test records;
- collect abnormal UI response clues with screenshots;
- optionally convert page/control records into MeterSphere import workbooks.

This skill is a testing assistant, not a defect oracle. Generated abnormal records must be reviewed manually before defects are filed.

## Repository Layout

```text
kmmom-ui-test-generator/
├─ README.md
├─ requirements.txt
├─ .gitignore
├─ skills/
│  └─ kmmom-ui-test-generator/
│     ├─ SKILL.md
│     ├─ agents/
│     ├─ assets/
│     ├─ references/
│     └─ scripts/
└─ docs/
   └─ usage-guide.md
```

## Install

Clone this repository:

```powershell
git clone https://github.com/<your-org-or-user>/kmmom-ui-test-generator.git
cd kmmom-ui-test-generator
```

Install Python dependencies:

```powershell
pip install -r requirements.txt
```

Copy the skill into the local Codex skills directory:

```powershell
Copy-Item -Recurse `
  .\skills\kmmom-ui-test-generator `
  "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator" `
  -Force
```

Or use the bundled installer:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

To install the skill and configure the local KMMOM password environment variable:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1 -SetPassword
```

The installer prompts for the password and saves it to the current Windows user's `KMMOM_PASSWORD` environment variable. The password is not stored in this repository.

Restart Codex after installing or updating the skill.

## Configure

Copy the example config and replace placeholders with local environment values:

```powershell
Copy-Item `
  .\skills\kmmom-ui-test-generator\assets\config.example.json `
  .\config.local.json
```

Important fields:

- `base_url`: KMMOM/MOM system URL.
- `debug_url`: Chrome CDP JSON endpoint, usually `http://127.0.0.1:9223/json/list`.
- `chrome_debug_port`: Chrome remote debugging port.
- `workspace`: local workspace where outputs are written.
- `module`: default module to traverse.
- `safe_mode`: keep this `true` unless a controlled test environment and cleanup plan exist.
- `password_env`: environment variable name used for login password, if scripted login is needed.

Do not commit `config.local.json`, passwords, screenshots, browser profiles, or execution outputs.

## Launch Chrome

Use the bundled launcher after confirming it is acceptable to open a controllable Chrome instance:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\launch_chrome.ps1" `
  -Url "http://192.168.30.69:40000" `
  -RemoteDebuggingPort 9223 `
  -UserDataDir "D:\Personal\Documents\KMMOM系统管理后台\chrome_kmmom_profile"
```

Log in to KMMOM/MOM in this browser before running page traversal.

## Generate UI Click Records

Example:

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\kmmom_ui_audit.py" `
  --base-url "http://192.168.30.69:40000" `
  --module "制造执行" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台" `
  --debug-url "http://127.0.0.1:9223/json/list" `
  --safe-mode
```

Outputs are written under:

```text
outputs/04_测试执行/02_执行记录/
```

## Convert To MeterSphere Cases

```powershell
$env:PYTHONIOENCODING='utf-8'
python "C:\Users\Administrator\.codex\skills\kmmom-ui-test-generator\scripts\generate_metersphere_cases.py" `
  --source "D:\Personal\Documents\KMMOM系统管理后台\outputs\04_测试执行\02_执行记录\KMMOM-制造执行模块-界面按钮点击测试表-YYYYMMDD.xlsx" `
  --workspace "D:\Personal\Documents\KMMOM系统管理后台"
```

Default granularity is page-level. Use `--granularity control` only when control-level cases are explicitly required.

## Safety Rules

Production-like environments must use safe mode.

Do not execute final actions such as:

- save;
- submit;
- delete;
- export or download;
- publish or issue;
- approve or reject;
- initialize;
- send messages.

For risky controls, record visibility, enabled state, expected result, dialog content, and risk reason. Do not confirm the final action unless the test environment and cleanup plan are explicitly approved.

## Review Abnormal Records

Abnormal records are clues, not final defect conclusions.

Review especially:

- login expiration;
- network or service failures;
- blank/error pages;
- API failures;
- permission prompts;
- business rule prompts that appear at unexpected times.

Do not file defects directly for:

- normal delete confirmations;
- popup residue;
- page content residue;
- coordinate or CDP timing errors;
- route changes that match normal page navigation.

See [docs/usage-guide.md](docs/usage-guide.md) for tester-facing usage guidance.
