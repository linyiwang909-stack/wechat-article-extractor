# 公众号文章提取器

输入微信公众号文章链接，提取标题、正文和配图，并生成可编辑的 Word 文档。

## 本地运行

```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```

打开 <http://127.0.0.1:8765/>。

## 部署到 Render

这个项目包含 Python 后端，不能直接用 GitHub Pages 部署。推荐使用 Render：

1. 在 GitHub 新建一个公开仓库，例如 `wechat-article-extractor`。
2. 在项目目录执行：

```bash
git init
git add app.py index.html requirements.txt .gitignore README.md
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<你的用户名>/wechat-article-extractor.git
git push -u origin main
```

3. 登录 <https://render.com/>，选择 **New +** -> **Web Service**，连接刚才的 GitHub 仓库。
4. 设置：
   - Runtime: `Python 3`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python app.py`
   - Plan: Free（测试使用即可）
5. 点击 **Create Web Service**，等待部署完成。

Render 会生成一个公开地址，例如 `https://wechat-article-extractor.onrender.com`，任何人都可以访问。

## 注意事项

- 免费实例一段时间没有访问后会休眠，第一次打开可能需要等待几十秒。
- 公众号页面、图片和微信接口可能存在访问限制；这不是 GitHub 或 Render 的限制。
- 生成的 Word 文件保存在运行实例的临时磁盘中，服务重启后不保证继续保留。
- 如果公开给所有人使用，建议后续增加请求频率限制，避免被大量调用。
