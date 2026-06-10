# KMMOM UI Test Output Rules

## Execution Records

Place UI execution records here:

```text
outputs/04_测试执行/02_执行记录/
```

Default filename:

```text
KMMOM-<模块名>模块-界面按钮点击测试表-YYYYMMDD.xlsx
```

Workbook sheets:

- `按钮点击测试表`
- `页面汇总`
- `测试说明`

Required main columns:

1. 序号
2. 模块
3. 菜单路径
4. 所属页面
5. 路由地址
6. 区域
7. 按钮/控件
8. 控件类型
9. 点击前置
10. 预期结果
11. 实测状态
12. 展开/弹窗/选项内容
13. 异常/弹窗提示
14. 是否执行点击
15. 风险说明

`异常/弹窗提示` must capture the exact visible prompt text observed during or immediately after each click/input interaction, especially:

- toast/message/notification text.
- alert banners or permission/login-expired prompts.
- modal/dialog/drawer title, body text, and footer button text.
- API failure, loading failure, blank page, error page, empty-data prompt, or other abnormal response.
- Whether the prompt blocks subsequent operation and whether it can be safely closed.

Leave this field empty only when no prompt, popup, or abnormal response is observed.

## MeterSphere Cases

Place MeterSphere import case files here:

```text
outputs/03_测试用例/03_MeterSphere导入/
```

Default filename:

```text
KMMOM-<模块名>模块-页面按钮控件点击测试用例-MeterSphere导入-YYYYMMDD.xlsx
```

## Naming

- Use the module display name from the KMMOM menu.
- Use the current date unless a source execution record date is available.
- Keep one module per workbook for reviewability.

## Verification

After generating Excel files, verify:

- File exists and is non-empty.
- Workbook opens with `openpyxl`.
- Sheet names are correct.
- Header row matches the expected template.
- Abnormal responses and popup prompts are recorded in `异常/弹窗提示` when observed.
- Summary count equals the number of main records or cases.
