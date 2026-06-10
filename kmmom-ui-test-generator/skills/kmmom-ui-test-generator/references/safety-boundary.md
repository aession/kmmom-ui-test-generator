# KMMOM UI Test Safety Boundary

## Default Mode

Default to `safe_mode=true` for KMMOM unless the user clearly states that the target is a disposable test environment and confirms data cleanup.

The goal of this skill is UI inspection and regression-assist testing. It is not allowed to mutate production business data by default.

## Safe Actions

These actions are normally safe:

- Open a page.
- Expand or collapse filters.
- Open dropdowns and read visible options.
- Open date pickers and close them.
- Type `TEST_NO_SAVE` into an input and clear it before leaving the field.
- Click query/search/reset when the page remains read-only.
- Open dialogs or drawers and close them with cancel, close icon, or Escape.
- Click table column settings, view settings, or page display controls.
- Record disabled controls without clicking them.

## Risky Actions

Do not execute final actions containing these texts:

- 保存
- 提交
- 确定
- 确 定
- 删除
- 导出
- 下载
- 初始化
- 发布
- 下发
- 通过
- 驳回
- 发送消息

Also treat icon-only buttons as risky when they are inside a footer/action area of a dialog and the action cannot be identified.

## Risk Recording

For skipped controls, record:

- Visible text or inferred control name.
- Area and page.
- Whether the control is enabled.
- Expected behavior.
- Actual status as `风险动作未点击` or `已识别`.
- Risk reason.

## When User Requests Full Operation

If the user asks to execute save/submit/delete/publish/downstream actions:

1. Confirm environment type.
2. Confirm test data ownership.
3. Confirm rollback or cleanup method.
4. Prefer using a dedicated low-privilege test account.
5. Record exact business impact in the output.
