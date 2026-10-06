# -*- coding: utf-8 -*-
"""通过 GitHub Contents API + Git Data API 上传项目（两步：种子提交 -> 全量提交）

用法：GITHUB_TOKEN=ghp_xxx python tests/api_push.py
Token 从环境变量读取，绝不写入本文件（GitHub push protection 会拦截含令牌的提交）。
"""
import base64, json, urllib.request, urllib.error, os

TOKEN = os.environ.get("GITHUB_TOKEN", "")
assert TOKEN, "请先设置环境变量 GITHUB_TOKEN"
REPO = "unfurlyang/ppt-study-hub"
API = "https://api.github.com"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 项目根目录

def req(method, path, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    r = urllib.request.Request(API + path, data=data, method=method)
    r.add_header("Authorization", "token " + TOKEN)
    r.add_header("Accept", "application/vnd.github+json")
    r.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(r, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))

def b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")

# --- 确定父提交：main 不存在则用 Contents API 播种 .gitignore ---
try:
    head = req("GET", "/repos/%s/git/ref/heads/main" % REPO)
    parent = head["object"]["sha"]
    print("main exists ->", parent)
except urllib.error.HTTPError as e:
    if e.code != 404:
        raise
    req("PUT", "/repos/%s/contents/.gitignore" % REPO, {
        "message": "chore: add .gitignore",
        "content": b64(os.path.join(ROOT, ".gitignore")),
        "branch": "main",
    })
    print("seed commit ok (.gitignore)")
    parent = req("GET", "/repos/%s/git/ref/heads/main" % REPO)["object"]["sha"]
    print("main ->", parent)

# --- Git Data API 上传全部文件，单个提交 ---
FILES = [
    ".gitignore",
    "index.html",
    "assets/app.js",
    "assets/config.js",
    "assets/style.css",
    "tests/make_sample.py",
    "tests/make_samples.py",
    "tests/make_big_txt.py",
    "tests/api_push.py",
    "tests/samples/sample.docx",
    "tests/samples/sample.pdf",
    "tests/samples/sample.txt",
    "tests/samples/sample.md",
    "tests/samples/big.txt",
    "tests/样例课件-牛顿第一定律.pptx",
    "tests/e2e-result.png",
]
tree = []
for f in FILES:
    content = b64(os.path.join(ROOT, *f.split("/")))
    try:
        b = req("POST", "/repos/%s/git/blobs" % REPO, {"content": content, "encoding": "base64"})
    except urllib.error.HTTPError as e:
        print("blob FAILED (%d): %s — 若为 422，请检查该文件是否含有令牌/密钥" % (e.code, f))
        raise
    tree.append({"path": f, "mode": "100644", "type": "blob", "sha": b["sha"]})
    print("blob ok:", f)

t = req("POST", "/repos/%s/git/trees" % REPO, {"tree": tree})
c = req("POST", "/repos/%s/git/commits" % REPO, {
    "message": "feat: 支持多格式课件解析（.pptx/.pdf/.docx/.txt/.md）+ 长文档自动分块",
    "tree": t["sha"],
    "parents": [parent],
})
r = req("PATCH", "/repos/%s/git/refs/heads/main" % REPO, {"sha": c["sha"], "force": False})
print("full commit:", c["sha"])
print("main ->", r["object"]["sha"])
