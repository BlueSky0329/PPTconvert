# 将项目推送到 GitHub（新建仓库）

在 **GitHub 网页** 上新建空仓库后，在本地项目根目录执行以下命令（将 `YOUR_USER` / `YOUR_REPO` 换成你的用户名与仓库名）。

## 1. 在 GitHub 创建仓库

1. 登录 [GitHub](https://github.com)，右上角 **+** → **New repository**。
2. **Repository name**：例如 `PPTconvert`。
3. 选择 **Public**（或 Private）。
4. **不要**勾选「Add a README」等初始化选项（保持空仓库，避免与本地首次推送冲突）。
5. 点击 **Create repository**，页面会显示推送命令，可对照下面步骤。

## 2. 本地首次推送（若尚未执行过 `git init`）

在项目根目录 `PPTconvert` 下打开终端（PowerShell 或 CMD）：

```powershell
cd 你的路径\PPTconvert

git init -b main
git add .
git commit -m "Initial commit: Word 转 PPT 工具"
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git push -u origin main
```

若仓库已存在且已做过首次提交，只需：

```powershell
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git branch -M main
git push -u origin main
```

## 3. 使用 SSH（可选）

若已配置 SSH 密钥，可将 `remote` 改为：

```text
git@github.com:YOUR_USER/YOUR_REPO.git
```

## 4. 常见问题

| 情况 | 处理 |
|------|------|
| 提示需要登录 | GitHub 已不支持账号密码推送 HTTPS，请使用 **Personal Access Token** 作为密码，或改用 **SSH**。 |
| `remote origin already exists` | `git remote remove origin` 后重新 `git remote add origin ...` |
| 推送被拒绝 | 若远程已有 README，可先 `git pull origin main --rebase` 再 `git push`。 |

完成后，在仓库 **Settings → General** 中可补充描述、勾选 **Topics**（如 `python`、`docx`、`pptx`、`education`）。
