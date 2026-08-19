# CODING Open API 速查（事项/缺陷/附件）

实测环境：e.coding.net，个人令牌认证 `Authorization: token <TOKEN>`。所有 Action 均为 POST 到 `https://e.coding.net/open-api`。

## 常用接口

| Action | 用途 | 关键参数 |
|---|---|---|
| DescribeCodingCurrentUser | 当前用户 | 无 |
| DescribeUserProjects | 用户项目列表 | UserId |
| DescribeProjectByName | 按名称查项目 | ProjectName → Response.Project.Id |
| DescribeProjectMembers | 项目成员 | ProjectId, PageNumber, PageSize |
| DescribeProjectIssueFieldList | 事项类型字段配置 | ProjectName, IssueType(DEFECT) |
| CreateIssue | 创建事项 | 见下 |
| ModifyIssue | 修改事项 | ProjectName, IssueCode, Name, FileIds |
| DescribeIssue | 事项详情 | ProjectName, IssueCode |
| DescribeIssueList | 事项列表 | ProjectName, IssueType |
| DescribePreSignUploadUrl | 文件预签名上传 | ProjectName, FileName, ContentType, FolderType, FolderId |
| CreateFile | 登记文件（拿文件ID） | AuthToken, StorageKey |
| DescribeIssueFileUrl | 附件下载地址 | ProjectName, **FileId 实为事项 Code** |

## CreateIssue 参数（DEFECT 实测）

```
Action, ProjectName, Type("DEFECT"), Name, Priority("0"~"3"字符串),
AssigneeId(处理人成员ID), Description, DefectTypeId(缺陷类型选项Value),
DueDate("YYYY-MM-DD"), CustomFieldValues([{Id, Content}])
```

必填校验顺序：AssigneeId → DefectTypeId → DueDate → 必填自定义字段。错误码：
`issue_assignee_required` / `issue_defect_type_required` / `issue_due_date_required` / `custom_field_required`。

## CustomFieldValues.Content 格式

- SELECT_SINGLE（单选）：选项 Value 字符串，如 "821994"
- SELECT_MULTI（多选）：选项 Value 字符串，如 "879094"（单值即可）
- SELECT_MEMBER_SINGLE / MULTI（成员）：成员 ID 字符串，如 "8351409"
- 处理人/关注人是系统参数（AssigneeId/WatcherIds），不进 CustomFieldValues

## 附件上传 4 步（2026-08-19 实跑验证）

1. `DescribePreSignUploadUrl`：FolderType=1（项目协同附件）/ 0（文件网盘）；返回 UploadLink、StorageKey、Headers（JSON 字符串，含 x-cos-security-token）、AuthToken
2. PUT UploadLink：headers = Content-Type + Content-Length + x-cos-security-token → 200
3. `CreateFile`：AuthToken + StorageKey → Response.Data.Id（文件 ID）
4. `ModifyIssue`：FileIds=[文件ID]

陷阱：
- `CreateAttachmentPrepareSignUrl` 返回的 AttachmentId 不能用于 FileIds（`issue_file_not_exist`）
- `DescribeIssueFileUrl` 的 FileId 参数实为事项 Code（查附件下载地址传 Code）
- AuthToken 与 StorageKey 必须来自同一次预签名，且预签名有时效（30分钟）

## 项目字段实测（BrainServicePlatform DEFECT，2026-08-19）

字段 ID 随项目变化，使用前必须查询，勿硬编码：

| 字段 | ID | 类型 | 必填 | 备注 |
|---|---|---|---|---|
| 处理人 | 35372238 | SELECT_MEMBER_SINGLE | 是 | AssigneeId 系统参数 |
| 问题归属人 | 38029242 | SELECT_MEMBER_MULTI | 是 | Content=成员ID |
| Bug归类 | 38411845 | SELECT_MULTI | 是 | web前端=879094 等 |
| 严重程度 | 38029238 | SELECT_SINGLE | 是 | 一般=821994 |
| 缺陷类型 | 35372239 | SELECT_SINGLE | 是 | DefectTypeId 系统参数，功能缺陷=35236068 |
| 问题来源 | 38029240 | SELECT_SINGLE | 是 | 测试环境=822004 |
| 引入该问题的项目 | 38029254 | SELECT_SINGLE | 是 | 服务区产品维护=907658 |
| 选择所属项目 | 38029247 | SELECT_MULTI | 是 | 公共项目=1091162 |
| 截止日期 | 35372242 | TEXT_DATE | 是 | DueDate 系统参数 |

## 参考

- 完整 spec.json（420 接口）：GitHub `yankeguo/coding-node-client` → spec.json
- apifox 镜像文档：`codingapi.apifox.cn/api-<id>.md`（接口 markdown 含完整 OpenAPI schema）
- CODING 官方文档：`help.coding.net/openapi`（SPA，锚点抓取困难，优先 apifox）
