# 旬度内贸排产优化模型

这是一个建筑用钢内贸旬度排产网页应用。

应用会根据外贸固定排产、产线生产天数、保供需求、规格比例、现金流等条件，生成内贸排产方案。

## 功能

- 在网页里直接录入数据，不需要上传 Excel。
- HRB500E 和 600兆帕按需求必须排满。
- 400兆帕钢筋在约束内追求现金流最优。
- 外贸排产天数如果超过产线生产天数，直接报错。
- 输出内贸排产方案、方案校验、400兆帕客户需求建议。

## 本地运行

先安装依赖：

```powershell
py -m pip install -r requirements.txt
```

启动网页：

```powershell
py web_app.py
```

打开：

```text
http://127.0.0.1:8765
```

如果只是自己电脑使用，也可以双击：

```text
start_schedule_web.cmd
```

## 上传 GitHub

请上传以下文件和文件夹：

- `web/`
- `schedule_model.py`
- `web_app.py`
- `requirements.txt`
- `README.md`
- `AGENT.md`
- `.gitignore`
- `Procfile`
- `render.yaml`
- `runtime.txt`

不要上传这些本地文件夹：

- `libs/`
- `.python-libs/`
- `.python-libs2/`
- `__pycache__/`

也不要上传真实业务 Excel 文件，避免数据泄露。

## Render 部署

推荐使用 Render 部署。

1. 把本项目上传到 GitHub。
2. 登录 Render。
3. 选择 New Web Service。
4. 连接这个 GitHub 仓库。
5. Render 会自动识别 `render.yaml`。
6. 部署完成后，Render 会给你一个公网网址。

如果手动填写配置：

- Build Command: `pip install -r requirements.txt`
- Start Command: `python web_app.py`
- Environment: Python

## 输入说明

网页里需要维护这些表：

- 旬度资源
- 生产日历
- 外贸排产明细
- 产品生产效率
- 现金流
- HRB500E及600兆帕需求
- 400兆帕规格比例
- 400兆帕客户需求预报

资源量如果小于 1000，模型按“万吨”理解；如果大于等于 1000，模型按“吨”理解。

## 输出说明

内贸排产方案字段：

- 旬度
- 牌号
- 规格
- 吨位
- 生产条线

400兆帕需求建议中的折扣：

```text
折扣 = 排产方案吨位 / 客户预报吨位
```
