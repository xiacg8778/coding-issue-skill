---
name: coding-issue-bug
description: 在 CODING DevOps（coding.net）创建 Bug 缺陷单，支持指定处理人/问题归属人/Bug归类，并可附带本地图片（支持多张）。当用户要求"在 CODING 建 bug/缺陷单/工作项"、"创建带图片的 bug"、"给 CODING 提缺陷"时使用。自动处理成员ID解析、必填字段补齐、附件上传（预签名+COS+登记+挂载）全链路。适合 AI/江苏高信/智慧服务区等 CODING 团队项目。
agent_created: true
---

# CODING 创建 Bug 单（含附件）

在 CODING DevOps 项目下创建缺陷单，支持处理人、问题归属人、Bug 归类、本地图片附件。

## 使用前提

- 需要 CODING token：`~/.workbuddy/mcp.json` 中 `mcpServers.coding-devops.env.CODING_TOKEN`（脚本自动读取）
- 目标项目在 `DescribeUserProjects` 中可见，成员在 `DescribeProjectMembers` 中可查
- **不要用 coding-devops MCP 的 create_issue**：它不传 AssigneeId，CODING API 报 `issue_assignee_required`，封装再读 `Response.Issue.Name` 崩溃（`Cannot read properties of undefined (reading 'Name')`）。一律直调 Open API。

## 快速路径（推荐）

直接运行脚本，一条命令完成全流程（成员解析 → 字段补齐 → 创建 → 附件挂载 → 验证）：

```bash
python ~/.workbuddy/skills/coding-issue-bug/scripts/create_bug.py \
  --project BrainServicePlatform \
  --title "这是一个test" \
  --description "啊啊啊哈哈哈" \
  --assignee 刘唯 \
  --owner 刘唯 \
  --category web前端 \
  --image C:/123.png
```

参数说明：`--project`/`--title`/`--description`/`--assignee` 必填；`--owner` 默认=处理人；`--category` 是 Bug 归类选项标题（不传取第一个选项）；`--image` 本地图片路径，**可多次传参支持多张**；`--priority` 0低1中2高3紧急（默认1）；`--due-date` 默认7天后。

多图示例：

```bash
python ~/.workbuddy/skills/coding-issue-bug/scripts/create_bug.py \
  --project BrainServicePlatform \
  --title "页面样式异常" \
  --description "复现步骤见附件" \
  --assignee 刘唯 \
  --category web前端 \
  --image C:/截图1.png \
  --image C:/截图2.png \
  --image C:/截图3.png
```

## 批量提单（QA-Team 测试报告 → Excel 补充 → 批量建单）

适用场景：qa-team 测试完成产出 `defect_ledger.json`（或缺陷草稿列表）后，导出 Excel 给用户补充项目/处理人等信息，再批量提单。

```bash
# 1. 导出 Excel 补充模板（QA 产出列灰底自动带出，用户补充列黄底标*必填）
python ~/.workbuddy/skills/coding-issue-bug/scripts/export_defects_excel.py <defect_ledger.json>

# 2. 用户在 Excel 黄底列补充：project*/assignee*/owner/category/images/priority/due_date/create*

# 3a. dry-run 校验（必做第一步，零写入）
python ~/.workbuddy/skills/coding-issue-bug/scripts/batch_create_bugs.py <补充后的.xlsx>

# 3b. 真实提单 + 结果回填 Excel（issue_code/issue_url/status 三列）
python ~/.workbuddy/skills/coding-issue-bug/scripts/batch_create_bugs.py <补充后的.xlsx> --execute --write-back
```

行为约定：`create` 列 no 的行跳过；必填缺失/图片路径不存在/项目不在 token 可见范围的行标 skipped 带原因，**单条失败不中断批次**，结束输出汇总。dry-run 预检走 `DescribeCodingCurrentUser` → `DescribeUserProjects`（⚠️ 必须带 UserId，无参报"团队成员不存在"）。事项页 URL 由 `resolve_team_host()` 动态解析团队域名（⚠️ `DescribeTeam` 的团队信息在 **`Data`** 键下，不是 `Team` 键——读错键拿到空会误判"接口不返回域名"而转向臆猜）。Excel 列结构由 `export_defects_excel.py` 的 COLUMNS 定义，两脚本配套使用。

### 铁律（批量提单，防重复踩坑）

1. **域名不猜**：团队域名只能从 `DescribeTeam` 的 `Data.TeamHost` 动态解析或用户实证获得，任何 `{project}.coding.net` 形态都是臆造，必 404。
2. **网络模糊失败 = 请求可能已生效**：CreateIssue 报网络错误（SSL 瞬断/超时）时，单据可能已建。重试前必须按标题查服务端（脚本已内置幂等②）；发现重复单用 `DeleteIssue` 清理，别让用户从 CODING 页面先发现。
3. **接口返回空先怀疑取错键**：下"接口不可用"结论前 dump 原始响应核对键名（本 skill 两次翻车均为取错键：DescribeUserProjects 缺 UserId、DescribeTeam 读 Team 而非 Data）。
4. **提单后必须核验**：`DescribeIssue` 逐单确认处理人/附件数/状态，Excel 回填 `--write-back`，defect_ledger 同步 tracker_id + submitted。

## 手动 API 路径（脚本不可用时）

所有请求：`POST https://e.coding.net/open-api`，Header `Authorization: token <CODING_TOKEN>`，Content-Type application/json。脚本本质也是走这些步骤，可对照。

### Step 1 解析成员 ID

```
Action: DescribeProjectMembers
{ProjectId: <项目ID>, PageNumber: 1, PageSize: 100}
→ Response.Data.ProjectMembers[].Id / Name
```
项目 ID 用 `DescribeProjectByName`（`ProjectName` → `Response.Project.Id`）获取。

### Step 2 查询 DEFECT 字段配置

```
Action: DescribeProjectIssueFieldList
{ProjectName: <项目>, IssueType: "DEFECT"}
→ Response.ProjectIssueFieldList[].IssueField{Id, Name, ComponentType, Options[{Title, Value}]}
```
关键字段（BrainServicePlatform 实测）：处理人(35372238, SELECT_MEMBER_SINGLE 必填)、问题归属人(38029242, SELECT_MEMBER_MULTI 必填)、Bug归类(38411845, SELECT_MULTI 必填, web前端=879094)、严重程度、缺陷类型、问题来源、引入该问题的项目、选择所属项目（后三者均必填）。**字段 ID 各项目不同，必须先查询再使用。**

### Step 3 创建 Bug 单（CreateIssue）

必填参数（逐层校验，缺什么报什么）：
- `AssigneeId`（处理人成员 ID）→ 缺了报 `issue_assignee_required`
- `DefectTypeId`（缺陷类型选项 Value，如功能缺陷=35236068）→ 缺了报 `issue_defect_type_required`
- `DueDate`（"YYYY-MM-DD" 字符串）→ 缺了报 `issue_due_date_required`
- `Priority`（"0"~"3" 字符串）
- `CustomFieldValues`：`[{Id: <字段ID>, Content: "<选项Value 或 成员ID>"}]`，覆盖所有必填自定义字段

```
Action: CreateIssue
{ProjectName, Type: "DEFECT", Name, Priority, AssigneeId, Description, DefectTypeId, DueDate, CustomFieldValues}
→ Response.Issue.Code（Bug 单编号）
```

### Step 4 上传图片附件（4 步，顺序不能错）

1. **预签名**：`DescribePreSignUploadUrl`，`{ProjectName, FileName, ContentType: "image/png", FolderType: 1, FolderId: 0}` → `Response.Data`：UploadLink / StorageKey / Headers(JSON, 含 x-cos-security-token) / AuthToken。**FolderType=1 是项目协同附件场景**，0 是文件网盘。
2. **上传**：PUT 到 UploadLink，Headers 带 `Content-Type` + `Content-Length` + Headers 里的 `x-cos-security-token`（响应码 200）。
3. **登记**：`CreateFile`，`{AuthToken, StorageKey}` → `Response.Data.Id` = 真实文件 ID。
4. **挂载**：`ModifyIssue`，`{ProjectName, IssueCode, Name, FileIds: [<文件ID>]}`。多张图片时逐个预签名→上传→登记，将全部文件 ID 收集进一个 `FileIds` 数组一次性挂载（已实测 2 张验证通过）。

⚠️ 陷阱：`CreateAttachmentPrepareSignUrl` 返回的 AttachmentId **不是** FileIds 要的 ID（会报 `issue_file_not_exist`）；`DescribeIssueFileUrl` 的 FileId 参数实为事项 Code（文档描述误导）。只有 `CreateFile` 登记的 ID 才能挂载。

## 故障排查

| 报错 | 原因 | 解法 |
|---|---|---|
| `Cannot read properties of undefined (reading 'Name')` | MCP 封装 bug | 直调 Open API，先 curl 复现真实错误 |
| `issue_assignee_required` | 缺 AssigneeId | 补处理人成员 ID |
| `issue_defect_type_required` | 缺 DefectTypeId | 补缺陷类型选项 Value |
| `issue_due_date_required` | 缺 DueDate | 补 "YYYY-MM-DD" |
| `issue_file_not_exist` | FileIds 传了附件预上传 ID | 走 CreateFile 登记拿真实文件 ID |
| `issue_project_file_not_exist` | CreateIssue 的 FileIds 也是文件网盘/登记 ID | 同上 |
| `auth_error`（CreateFile） | AuthToken/StorageKey 不匹配或已过期 | 重新走 Step 4 全流程 |
| `团队成员不存在`（DescribeUserProjects） | 未传 UserId | 先 `DescribeCodingCurrentUser` 取 Id 再查询 |
| 事项页 URL 404 | 团队域名不可臆造（`{project}.coding.net` 不存在） | `DescribeProjectByName`→TeamId→`DescribeTeam` 读 **`Data.TeamHost`**；格式 `{TeamHost}/p/{project}/bug-tracking/issues/{code}` |
| `SSL: UNEXPECTED_EOF_WHILE_READING` | e.coding.net 网络瞬断 | call_api 已内置 3 次退避重试；仍失败按行跳过不中断批次 |
| 批量重跑建出重复单 | CreateIssue 已送达但响应丢失 → 客户端误判 failed | 重提前按标题查 `DescribeIssueList` 认领已有单（脚本幂等②已内置）；重复废单用 `DeleteIssue` 清理 |

## 注意

- 公司截图禁传外部公开托管（Mr夏红线）；CODING COS 内部通道合规。
- 创建后必须 `DescribeIssue` 验证（处理人/归属人/Bug归类/附件数），再向用户报告。
- 完整接口 spec 参考：GitHub `yankeguo/coding-node-client` 的 `spec.json`（420 个接口，含全部字段定义）。

## Resources

- `scripts/create_bug.py`：一键创建脚本（推荐，含成员解析/字段补齐/附件上传）
- `scripts/export_defects_excel.py`：缺陷台账 → Excel 补充模板导出（qa-team 联动入口）
- `scripts/batch_create_bugs.py`：Excel 驱动批量提单（默认 dry-run，--execute 真实创建，--write-back 回填）
- `references/api_reference.md`：接口与字段细节速查
