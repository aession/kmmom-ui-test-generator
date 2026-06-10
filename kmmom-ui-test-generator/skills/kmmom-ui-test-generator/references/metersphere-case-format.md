# MeterSphere Case Format

## Columns

Use these columns exactly, in this order:

1. 用例名称
2. 所属模块
3. 标签
4. 前置条件
5. 步骤描述
6. 预期结果
7. 编辑模式
8. 备注
9. 用例状态
10. 责任人
11. 用例等级

## Default Values

- `编辑模式`: `TEST`
- `用例状态`: `未开始`
- `责任人`: blank unless the user provides one
- `标签`: `KMMOM,模块名,页面按钮控件点击`

## Row Expansion

One test case may have multiple steps. Expand it into multiple Excel rows:

- First row: fill all shared fields plus step 1 and expected result 1.
- Following rows: fill only `步骤描述` and `预期结果`.
- Leave other fields blank on following rows.

## Case Granularity

Default:

- One page/interface = one MeterSphere test case.
- Each visible button/control/filter/dropdown/dialog entry/table control on the page = a step in that case.

Optional:

- One control = one MeterSphere test case only when the user explicitly asks for fine-grained traceability or when a single control has complex independent rules.

## Case Naming

Default page-level case name:

```text
MES-UI-PAGE-<序号> 验证【<页面>】页面按钮/控件自动化点击巡检
```

Use module-specific prefixes when helpful, for example:

```text
MES-EXEC-UI-PAGE-001
```

## Priority

Suggested priority mapping:

- `P1`: toolbar buttons, filters, page condition controls, risky final-action controls that must be reviewed.
- `P2`: dialog internal controls, table controls, dropdown options, ordinary UI interactions.
- `P3`: public floating controls or low-risk auxiliary controls.

## Remark Content

The `备注` field should include:

- Source execution record.
- Scenario type.
- Menu path.
- Route address.
- Page control record count.
- Executed click/expand/input count.
- Safe-identification or risk-skip count.
- Statement that the case is page-level and controls are expanded as steps.
