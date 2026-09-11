# Excel 导航插件：RustFS 上传与天眼路由接入

## 当前状态与分工

- 存储桶：`excel-navigation`，用户已创建并保存匿名只读策略，版本控制已开启。
- 上传账号：用户已确认创建 专用上传账号/策略 及同名策略；本机凭据已通过真实上传与权限检查。
- 本机项目根目录 `.env` 已配置六项 `TYAPP_S3_*` 参数，已加入 Git 忽略规则；不在日志中输出凭据。
- 天眼侧负责路由配置与部署。本次未修改服务器配置；此前本地两处路由草稿已撤回。
- ExcelPane 侧负责打包、上传及发布验证。1.0.0.2 已上传至 GitHub/RustFS；1.0.1.0 改为 EXE 内置 MSI，目前为本地待验收版本，尚未上传。

本机复测（2026-09-11 01:28）：`python tests/RustFSUploadCheck.py` 通过。签名列举、上传与覆盖、签名下载、匿名下载及 SHA-256 校验均通过，包含中文、空格、加号、百分号的对象路径。匿名列举和上传、跨桶列举、读取桶策略、删除不存在的测试对象均返回 403。测试仅保留 `diagnostics/upload-check-8d7bf50d2876422c8a2dd4268d3a2219/` 下的小文本对象及覆盖版本，未删除任何对象。

结果保存在 `work/rustfs/upload-check.json`。随后 1.0.0.2 的 `.vsto` 及发布文件已上传并校验。MSI 版沿用同一桶、账号和域名，新增 `latest.xml` 入口，发布时再上传；真实安装和升级仍待独立环境验收。以下路由要求保留作为维护约定。

## 上传账号权限

在 RustFS 控制台“策略”中创建 专用上传账号/策略，再创建同名用户并只关联该策略，不加入管理员组、不绑定 `readwrite` 或 `consoleAdmin`。

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": ["s3:GetBucketLocation", "s3:ListBucket", "s3:ListBucketMultipartUploads"],
      "Resource": ["arn:aws:s3:::excel-navigation"]
    },
    {
      "Effect": "Allow",
      "Action": ["s3:GetObject", "s3:PutObject", "s3:AbortMultipartUpload", "s3:ListMultipartUploadParts"],
      "Resource": ["arn:aws:s3:::excel-navigation/*"]
    }
  ]
}
```

允许列举本桶、上传/覆盖、下载校验及分片上传；不授予删除对象、修改桶策略、管理用户或其他桶的权限。匿名下载仍由桶策略决定。发布凭据仅供发布端使用，不进入插件、安装包、Git 或更新日志。

## 给天眼侧的路由要求

沿用现有域名，不需要新域名。期望接入参数：

| 项目 | 期望值 |
| --- | --- |
| S3 Endpoint | `https://oneview.jiarui.net.cn` |
| Bucket | `excel-navigation` |
| 寻址方式 / Region | Path style / `us-east-1` |
| 插件更新目录 | `https://oneview.jiarui.net.cn/excel-navigation/` |
| MSI 版更新入口 | `https://oneview.jiarui.net.cn/excel-navigation/latest.xml` |
| 安装包示例 | `https://oneview.jiarui.net.cn/excel-navigation/releases/1.0.1.0/ExcelNavigator-Setup-1.0.1.0.exe` |

天眼侧已完成桶路由接入；迁移 MSI 不需要再添加桶或域名。维护时，`tianyan/compose.object-storage.yaml` 的公共 S3 Traefik 路由及 `tianyan/object-storage/nginx.conf.template` 的 8080 网关应继续匹配 `excel-navigation`，路径表达式为：

```text
^/(tyobj-[a-z0-9]([a-z0-9-]{0,55}[a-z0-9])?|excel-navigation)(/|$)
```

- 保留现有 Host、可信代理及其他路由约束；不要开放所有桶或修改控制台 SSO 路由。
- 请求转发到现有 RustFS S3 服务 `tianyan-rustfs:9000`，不是控制台 9001。保持完整 Bucket/Object 路径、原始编码、查询参数及 SigV4 Authorization；不增加 StripPrefix 或重定向。
- 保留现有已验证的公网 Host 恢复方式，否则上传签名可能不匹配。支持带空格的 `Application Files/...` 和中文文件名。
- 该桶的公开对象读取不依赖 OneView 登录 Cookie。匿名只读由 RustFS 桶策略控制，写入必须验签；S3 错误应原样返回，不能回落到 OneView 首页。
- 新 `latest.xml` 上传为 `application/xml`；客户端按正文解析，即使网关按二进制下载也可工作。旧 `.vsto` 继续保留 `application/x-ms-vsto`，不删除或改写旧入口。
- 根更新清单和安装入口避免长期缓存，可沿用 `Cache-Control: no-store`。保留下载附件、`nosniff`、CSP 和剥离 Cookie 等既有防护。

## 接通后的验收

1. 使用专用账号经公网 S3 Endpoint 上传一个不含敏感信息的小测试文件；下载并核对 SHA-256。
2. 不带 Cookie、签名或登录状态访问同一文件，确认正文一致；缺失对象返回 S3 错误，而不是 HTML 首页。
3. 验证本桶列举成功，其他桶列举和桶策略读取被拒绝；匿名上传被拒绝。不用删除真实对象测试删除权限。
4. MSI 版先上传 `releases/版本号/` 内 EXE、SHA-256、两种位数 MSI、安装说明、签名清单和公钥并校验，最后覆盖根 `latest.xml`；保留旧版本及所有旧 ClickOnce 文件，签名后不得改写。
5. 独立测试环境先卸载旧 ClickOnce 或开发版注册，再安装 MSI 版；后续从菜单检查更新、下载新版 EXE、关闭 Excel 后运行升级。构建和 HTTP 下载通过不等于 Excel 升级通过。

完整 GitHub 与 RustFS 发布顺序见项目 README“发布流程”。
