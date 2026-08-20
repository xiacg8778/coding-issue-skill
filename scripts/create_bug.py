#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CODING DevOps 创建 Bug 单（含附件）脚本
基于 2026-08-19 实跑验证的完整流程封装。

用法:
  python create_bug.py --project BrainServicePlatform --title "标题" --description "描述" \
      --assignee 刘唯 --owner 刘唯 --category web前端 --image C:/123.png \
      --image C:/124.png --image C:/125.png [--priority 1] [--due-date 2026-08-26]

参数:
  --project      项目名称（必填）
  --title        标题（必填）
  --description  描述/内容（必填）
  --assignee     处理人姓名（必填，按项目成员姓名匹配）
  --owner        问题归属人姓名（可选，默认同处理人）
  --category     Bug 归类选项标题（可选，如 web前端；不传则取第一个选项）
  --image        图片路径（可选，本地绝对路径；可多次传参支持多张）
  --priority     优先级 0低 1中 2高 3紧急（可选，默认 1）
  --due-date     截止日期 YYYY-MM-DD（可选，默认 7 天后）
  --token        CODING token（可选，默认读环境变量 CODING_TOKEN 或 ~/.workbuddy/mcp.json）

token 获取: ~/.workbuddy/mcp.json 中 coding-devops 的 env.CODING_TOKEN
"""
import argparse
import datetime
import json
import os
import sys
import urllib.request
import urllib.error

API_URL = "https://e.coding.net/open-api"


def get_token(arg_token):
    if arg_token:
        return arg_token
    env = os.environ.get("CODING_TOKEN")
    if env:
        return env
    mcp_path = os.path.expanduser("~/.workbuddy/mcp.json")
    try:
        with open(mcp_path, encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg["mcpServers"]["coding-devops"]["env"]["CODING_TOKEN"]
    except Exception:
        sys.exit("错误: 无法获取 CODING token（--token / CODING_TOKEN / mcp.json）")


def call_api(token, payload, verbose=False):
    """调用 CODING Open API，返回 Response 字典；出错则打印错误并退出
    网络类异常（SSL 瞬断/超时/连接重置）自动退避重试 3 次（1s/2s/4s）"""
    import time
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(API_URL, data=data, method="POST")
    req.add_header("Authorization", f"token {token}")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    last_net_err = None
    for attempt in range(4):  # 1 次原始 + 3 次重试
        try:
            resp = urllib.request.urlopen(req, timeout=60)
            d = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as e:
            sys.exit(f"HTTP {e.code}: {e.read().decode('utf-8', 'ignore')[:500]}")
        except urllib.error.URLError as e:
            last_net_err = e
            if attempt < 3:
                wait = 2 ** attempt
                print(f"    [网络重试] {attempt + 1}/3: {str(e)[:80]}，{wait}s 后重试")
                time.sleep(wait)
                # urllib Request 对象 urlopen 后不可复用，重建
                req = urllib.request.Request(API_URL, data=data, method="POST")
                req.add_header("Authorization", f"token {token}")
                req.add_header("Content-Type", "application/json")
                req.add_header("Accept", "application/json")
    else:
        sys.exit(f"网络错误（重试后仍失败）: {last_net_err}")
    response = d.get("Response", {})
    if "Error" in response:
        err = response["Error"]
        sys.exit(f"API 错误 [{err.get('Code', '')}]: {err.get('Message', '')}")
    return response


def find_member(token, project_id, name):
    """按姓名在项目成员中查找用户 ID"""
    members = []
    page = 1
    while True:
        resp = call_api(token, {
            "Action": "DescribeProjectMembers",
            "ProjectId": project_id,
            "PageNumber": page,
            "PageSize": 100,
        })
        data = resp.get("Data", {})
        batch = data.get("ProjectMembers", [])
        members.extend(batch)
        if len(members) >= data.get("TotalCount", 0) or not batch:
            break
        page += 1
    for m in members:
        if name in (m.get("Name") or ""):
            return m["Id"]
    names = "、".join(m.get("Name", "") for m in members)
    sys.exit(f"错误: 项目成员中找不到「{name}」。现有成员: {names[:300]}")


def get_defect_fields(token, project_name):
    """查询 DEFECT 类型属性字段，返回 {名称: {field_id, component_type, required, options}}"""
    resp = call_api(token, {
        "Action": "DescribeProjectIssueFieldList",
        "ProjectName": project_name,
        "IssueType": "DEFECT",
    })
    fields = {}
    for item in resp.get("ProjectIssueFieldList", []):
        f = item.get("IssueField", {})
        fields[f.get("Name")] = {
            "field_id": f.get("Id"),
            "component_type": f.get("ComponentType"),
            "required": item.get("Required") == "True" or item.get("Required") is True,
            "options": {o.get("Title"): o.get("Value") for o in (f.get("Options") or [])},
        }
    return fields


def pick_option(field, prefer=None):
    """从字段选项里挑值 ID：优先 prefer（标题），否则取第一个"""
    opts = field["options"]
    if not opts:
        return None
    if prefer and prefer in opts:
        return opts[prefer]
    return next(iter(opts.values()))


def upload_attachment(token, project_name, image_path):
    """完整附件上传流程: 预签名 -> 上传 COS -> CreateFile 登记 -> 返回文件 ID"""
    fname = os.path.basename(image_path)
    fsize = os.path.getsize(image_path)
    # 1. 获取预签名上传信息（FolderType=1 表示项目协同附件）
    resp = call_api(token, {
        "Action": "DescribePreSignUploadUrl",
        "ProjectName": project_name,
        "FileName": fname,
        "ContentType": "image/png" if fname.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")) else "application/octet-stream",
        "FolderType": 1,
        "FolderId": 0,
    })
    data = resp["Data"]
    upload_link = data["UploadLink"]
    storage_key = data["StorageKey"]
    auth_token = data["AuthToken"]
    cos_headers = json.loads(data.get("Headers", "{}"))
    # 2. 上传文件到 COS（带 x-cos-security-token）
    file_data = open(image_path, "rb").read()
    req = urllib.request.Request(upload_link, data=file_data, method="PUT")
    req.add_header("Content-Type", "image/png" if fname.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")) else "application/octet-stream")
    req.add_header("Content-Length", str(len(file_data)))
    for k, v in cos_headers.items():
        req.add_header(k, v)
    try:
        urllib.request.urlopen(req, timeout=120)
    except urllib.error.HTTPError as e:
        sys.exit(f"图片上传失败 HTTP {e.code}: {e.read()[:300]}")
    # 3. CreateFile 登记，拿到真实文件 ID
    resp = call_api(token, {"Action": "CreateFile", "AuthToken": auth_token, "StorageKey": storage_key})
    return resp["Data"]["Id"]


def main():
    parser = argparse.ArgumentParser(description="CODING 创建 Bug 单（支持图片附件）")
    parser.add_argument("--project", required=True, help="项目名称")
    parser.add_argument("--title", required=True, help="标题")
    parser.add_argument("--description", required=True, help="内容/描述")
    parser.add_argument("--assignee", required=True, help="处理人姓名")
    parser.add_argument("--owner", default=None, help="问题归属人姓名（默认=处理人）")
    parser.add_argument("--category", default=None, help="Bug 归类选项标题（如 web前端）")
    parser.add_argument("--image", action="append", default=None, help="图片路径（可多次传参，支持多张）")
    parser.add_argument("--priority", default="1", help="优先级 0低/1中/2高/3紧急")
    parser.add_argument("--due-date", default=None, help="截止日期 YYYY-MM-DD")
    parser.add_argument("--token", default=None, help="CODING token")
    args = parser.parse_args()

    token = get_token(args.token)
    owner = args.owner or args.assignee
    due_date = args.due_date or (datetime.date.today() + datetime.timedelta(days=7)).isoformat()

    # 1. 查项目成员，解析处理人/归属人 ID
    resp = call_api(token, {"Action": "DescribeProjectByName", "ProjectName": args.project})
    project_id = resp.get("Project", {}).get("Id")
    assignee_id = find_member(token, project_id, args.assignee)
    owner_id = find_member(token, project_id, owner)
    print(f"[1/4] 成员解析完成: {args.assignee}={assignee_id}, {owner}={owner_id}")

    # 2. 查询 DEFECT 字段配置，构建必填自定义字段
    fields = get_defect_fields(token, args.project)
    custom_values = []
    # 处理人（系统字段）
    assignee_id = assignee_id
    # 缺陷类型（系统参数）
    defect_type_id = None
    if "缺陷类型" in fields:
        defect_type_id = int(pick_option(fields["缺陷类型"], prefer="功能缺陷"))
    # Bug 归类（自定义，用户指定）
    category_val = None
    if "Bug归类" in fields:
        category_val = pick_option(fields["Bug归类"], prefer=args.category)
        if args.category and category_val is None:
            sys.exit(f"错误: Bug 归类没有「{args.category}」选项。可选: {list(fields['Bug归类']['options'])}")
    # 问题归属人（自定义，成员 ID）
    custom_values.append({"Id": fields["问题归属人"]["field_id"], "Content": str(owner_id)})
    if category_val:
        custom_values.append({"Id": fields["Bug归类"]["field_id"], "Content": str(category_val)})
    # 其他必填自定义字段，未指定时取默认（跳过系统参数字段：处理人/缺陷类型/优先级/截止日期）
    for fname, f in fields.items():
        if fname in ("处理人", "缺陷类型", "问题归属人", "Bug归类", "关注人", "优先级", "截止日期", "开始日期", "开发截止日期", "进度"):
            continue
        if f["required"] and f["component_type"] in ("SELECT_SINGLE", "SELECT_MULTI"):
            val = pick_option(f, prefer={"严重程度": "一般", "问题来源": "测试环境"}.get(fname))
            if val:
                custom_values.append({"Id": f["field_id"], "Content": str(val)})
                print(f"    必填字段 {fname} = {val}")

    # 3. 创建 Bug 单
    payload = {
        "Action": "CreateIssue",
        "ProjectName": args.project,
        "Type": "DEFECT",
        "Name": args.title,
        "Priority": args.priority,
        "AssigneeId": assignee_id,
        "Description": args.description,
        "DueDate": due_date,
        "CustomFieldValues": custom_values,
    }
    if defect_type_id:
        payload["DefectTypeId"] = defect_type_id
    resp = call_api(token, payload)
    issue = resp["Issue"]
    issue_code = issue["Code"]
    print(f"[2/4] Bug 单创建成功: Code={issue_code} 状态={issue.get('IssueStatusName')}")

    # 4. 上传并挂载附件（支持多张）
    file_ids = []
    if args.image:
        for img in args.image:
            if not os.path.exists(img):
                print(f"[警告] 图片不存在，跳过: {img}")
                continue
            fid = upload_attachment(token, args.project, img)
            file_ids.append(fid)
            print(f"    附件已上传: {os.path.basename(img)} -> file_id={fid}")
        if file_ids:
            call_api(token, {
                "Action": "ModifyIssue",
                "ProjectName": args.project,
                "IssueCode": issue_code,
                "Name": args.title,
                "FileIds": file_ids,
            })
            print(f"[3/4] 附件已挂载: {len(file_ids)} 个")

    # 5. 验证
    final = call_api(token, {"Action": "DescribeIssue", "ProjectName": args.project, "IssueCode": issue_code})["Issue"]
    print(f"[4/4] 验证完成: 标题={final['Name']} 处理人={final['Assignee']['Name']} "
          f"附件数={len(final.get('Files', []))} 附件={[f['Name'] for f in final.get('Files', [])]}")
    print(f"\n✅ Bug 单创建完成: Code={issue_code}（项目 {args.project}）")


if __name__ == "__main__":
    main()
