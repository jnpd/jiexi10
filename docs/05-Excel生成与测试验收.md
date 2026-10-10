# 05｜Excel生成规范、测试矩阵与验收

> 验收目标：**Excel正确不等于产品识别正确**。必须同时核查“源文件范围、产品数、每张Sheet的字段数、字段证据与参数适用条件”。

## 一、导出规则（强制）

1. **一份客户源规格书 = 一个xlsx**，多份绝不混在一个源文件工作簿；订单附件不另算独立客户规格书，除非它本身被声明为需解析的源文件。
2. **一个被选中的ProductItem = 一个Sheet**；该工作簿不增加“统计/封面/汇总”Sheet。产品总汇总输出工作簿外的 Markdown / JSON。
3. 每张Sheet独立且完整：文件名/客户、产品8段名称、相数、额定高压、额定低压、额定容量、材质、频率等写在表头；正文写该次模板**全部字段、原始顺序、原单位**。
4. 正文建议列：`序号 | 分类 | 参数字段 | 模板单位 | 本产品参数值 | 提取状态 | 适用性/数值语义 | 原文出处`。
5. 参数表只有百分比损耗允差字段时，W值不塞进百分比格；可以在表头下的“独立损耗信息”块标`noLoadLossW/loadLossW`及其“规范上限”语义，或另建元数据列，但**不能增加第32个模板字段冒充来自模板**。
6. 每个Sheet的参数来源必须独立可读，不能引用别的Sheet的公式来取值。
7. 文件名保留源文件基名（路径非法字符归一化），再加 `_逐产品参数.xlsx`；避免不同同名文件发生覆盖，可在存储key里加 documentId。

## 二、Sheet命名

- 仅一个大类时用 `50kVA`、`75kVA`、`100kVA` 等。
- 同一工作簿包含多个相数或高低压等级时：`3P-11kV-0.4kV-75kVA`（如长度允许）。
- Excel Sheet名称最大31字符，包含 `[]:*?/\\` 等字符要替换；发生重名自动加确定性后缀 `-2`、`-3`，同时记录完整 productKey。
- 显示名称、参数值不作为 Sheet 唯一键；即使名称相同，产品键和源文件ID也应能定位。

## 三、建议 Excel 页面草图

```text
Sheet 标签：[75kVA]
A1:H1  铜-75kVA-11kV-0.4kV-空载200W-负载910W-4%-50Hz
A2:H2  客户：EDM  | 原文件：Attachment 1.1...pdf
A3:H3  相数：3P | 一次：11kV | 二次：400/230V（额定线电压0.4kV）
A4:H4  提取口径：STANDARD_SERIES/EXPLICIT_SPEC | 审核状态
A5:H5  额外指标：空载损耗≤200W、负载损耗≤910W（仅此75kVA，规范上限）
A7:H7  序号 | 分类 | 字段 | 单位 | 参数内容 | 状态 | 语义 | 原文证据
A8:H8  1    | 产品参数 | 标准 | — | IEC 60076 | FOUND_SHARED | 规范要求 | p5 §2
...
A27:H27 20 | ... | 空载上限 | % | +10%（测试允差） | FOUND_SHARED | TEST_TOLERANCE | p13表3.2
...
（按模板字段总数继续，不遗漏）
```

**显示与导出分离**：数据库 `valueNormalized`保存准确数值，Excel `valueDisplay`展示 `≤4%`、`+10%`、`未提供`等语义；产品名称可使用客户规定的“4%”简写，但参数行不省略`≤`。

## 四、Excel生成器伪代码

```python
def export_one_document(document, selected_products, template_fields, decisions, workbook):
    assert all(p["documentId"] == document["id"] for p in selected_products)
    assert len({p["productKey"] for p in selected_products}) == len(selected_products)

    for product in selected_products:
        sheet = workbook.add_worksheet(make_unique_sheet_name(product))
        write_product_header(sheet, product, document)
        rows = decisions[product["productId"]]
        assert {r["fieldCode"] for r in rows} == {f["code"] for f in template_fields}

        for index, f in enumerate(template_fields):
            row = next(r for r in rows if r["fieldCode"] == f["code"])
            write_parameter_row(sheet, index, f, row)

    assert workbook.sheet_count == len(selected_products)
    # 保存后要重新打开文件验证，而不是只信任内存中的对象
    return workbook
```

正式实现可以用 openpyxl 或已验证可用的 Excel生成脚本，重点是将生成器限制成**只消费已校验结构化JSON**，不让它从PDF猜字段。

## 五、真实样本测试用例（EDM 11kV/22kV）

| 编号 | 输入/场景 | 预期 | 失败例 |
|---|---|---|---|
| T01 | 11kV PDF、`STANDARD_SERIES` | 一份xlsx，9 Sheet | 只有1 Sheet |
| T02 | 22kV PDF、`STANDARD_SERIES` | 一份xlsx，9 Sheet | 与11kV合成一份 |
| T03 | 两份同时上传、`STANDARD_SERIES` | 两个xlsx，总18 Sheet | 错将两份相同容量去重 |
| T04 | 两份均按Scope识别明确规格 | 两份各至少有75kVA“明确规格”声明；11kV GTP亦指定75 | 将所有9个认作已采购 |
| T05 | 两份PDF没有订单 | `ORDER_CONFIRMED`无可确认清单、提示需订单 | 声称采购18型号 |
| T06 | 11kV 75kVA | 空载≤200W、负载≤910W，名字简写200/910 | 记录为厂家实测/保证值 |
| T07 | 11kV 50/100/160kVA | 空载待定、负载待定 | 复制75kVA损耗 |
| T08 | 11kV 315/500/630kVA | 阻抗限值≤4/≤5.5/≤6% | 全部4% |
| T09 | 22kV Um=24kV | 高压产品键保持22kV | 错误生成24kV型号 |
| T10 | LV 400/230V | 低压键0.4kV，仅一大类 | 同时生成0.23kV |
| T11 | 模板当前31字段 | 每个Sheet恰31模板字段 | 只输出有值字段 |
| T12 | 模板改为32字段 | 所有Sheet自动有32字段 | 硬编码31 |
| T13 | 同文件人工批注与正文冲突 | 标记CONFLICT并可查看原文 | 静默择一 |
| T14 | 文档正文“Vendor to state” | WAIT_VENDOR，不猜值 | 自动填行业标准值 |
| T15 | 一份文档出现1P和3P独立规格 | 两个大类分别展开 | 合并为一类 |
| T16 | 文件出现11/22kV与若干容量但没有配对关系 | 产生未确认组合、不直接相乘 | 擅自展开所有组合 |
| T17 | 解析产品引用不同documentId证据 | 阻断跨文档串值 | 自动从相似文件补齐 |
| T18 | PDF提取不到文本但页面有扫描文字 | 标记待视觉解析并进行局部处理 | 认为没有产品 |
| T19 | Excel导出任务重试 | 重试后规格数一致，不产生重复Sheet | 11kV生成18Sheet |
| T20 | 人工确认后修改值 | 产生新快照，旧快照不可变 | 老Excel内容被静默替换 |
| T21 | 文件同名不同内容 | documentId与hash不同，分别保存 | 覆盖上传文件 |
| T22 | Sheet名冲突或超过31字符 | 确定性截断/后缀，表头保留全名 | 导出失败或覆盖Sheet |
| T23 | 两份文档中相同容量组合 | 每个文档独立统计，另表可给全局去重 | 全局去重导致原文消失 |
| T24 | 数字百分比和W混淆 | 单位校验阻断/标注冲突 | +10%填成10W |
| T25 | 多文档批处理部分失败 | 成功文件状态可见，失败文件带原因可单独重跑 | 一条失败造成全批静默丢失 |

测试**不仅断言数量**，还要按键逐项对比来源、条件及值。对于本样本，160kVA虽不在模板参考三相标准容量清单，但必须保留并记录“客户值不在模板参考集合”；模板参考值不是否定性业务约束。

## 六、自动断言示例

```python
def assert_export(document, products, template, workbook):
    assert workbook.sheet_count == len(products)
    keys = [p["productKey"] for p in products]
    assert len(keys) == len(set(keys))

    for sheet, product in zip(workbook.sheets, products):
        assert sheet.header["productKey"] == product["productKey"]
        assert sheet.template_field_codes == [x["code"] for x in template["fields"]]
        assert sheet.template_units == [x["unit"] for x in template["fields"]]
        assert all(e["documentId"] == document["id"]
                   for e in sheet.primary_document_evidence)
```

注意人工审核内容、绑定的RFQ/订单作为另外的来源对象保存，不应通过改造`documentId`来伪装同一个源文件。

## 七、交付验收清单

- [ ] 页面可显示每份文件的系列数、具体规格数、已确认范围及待解决问题。
- [ ] 同一源PDF生成自己的xlsx，sheet数和“所选导出范围”的具体产品数一致。
- [ ] 31字段只是一份模板当前字段数；任意新模板都能动态产出。
- [ ] 所有字段保留模板名称、单位、顺序，缺值不会删除字段。
- [ ] 各产品W损耗适用严格限定；阻抗限值/允差/损耗W语义不混淆。
- [ ] 用户可定位证据，所有人工修改有记录，冲突未审不得正式确认。
- [ ] 导出可重复且幂等，重新读取.xlsx后的行数、产品名、值一致。
- [ ] 统计区分标准系列/明确规格/订单确认，绝不把技术系列数当采购台数。
