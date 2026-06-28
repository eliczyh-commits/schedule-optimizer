# GitHub 上传说明

我已经把项目整理成可以上传 GitHub 的结构。

## 最简单做法

上传这个压缩包里的内容：

```text
排产模型-GitHub上传包.zip
```

如果你看到的是文件夹，就上传：

```text
GitHub上传包
```

## 上传后怎么部署

推荐使用 Render。

1. 打开 https://render.com
2. 注册或登录。
3. 点击 New。
4. 选择 Web Service。
5. 连接你的 GitHub 仓库。
6. 选择刚上传的排产模型项目。
7. 如果 Render 自动识别配置，直接点 Deploy。

如果需要手动填写：

- Build Command: `pip install -r requirements.txt`
- Start Command: `python web_app.py`

部署完成后，Render 会给你一个网址。以后你和同事打开这个网址就能使用网页应用。

## 不要上传的内容

不要上传真实业务 Excel 文件。

不要上传这些目录：

- `libs`
- `.python-libs`
- `.python-libs2`
- `__pycache__`
