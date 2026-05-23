# HCI 巡检整改看板生成器

把深信服 HCI 原始巡检报告 `.docx` 上传到网页，自动生成客户展示版“HCI 巡检整改看板”Word 文件。

## 功能

- 自动提取巡检得分、异常项、告警项、设备 IP、备份覆盖率。
- 自动识别原始报告中的异常/告警风险项。
- 输出横版 Word 看板，包含：
  - 首页 4 个关键指标
  - 6 条以内核心整改项
  - P0/P1 优先级
  - 整改路线图
  - 风险明细备查

## 本地运行

```bash
python3 -m pip install -r requirements.txt
vercel dev
```

打开 `http://localhost:3000`，上传原始巡检报告即可生成。

## 部署到 Vercel

```bash
vercel
vercel --prod
```

部署后访问 Vercel 分配的域名，上传 `.docx` 文件即可下载生成后的整改看板。

## 项目结构

```text
api/convert.py      后端接口：解析巡检报告并生成 Word
public/index.html   上传页面
requirements.txt    Python 依赖
vercel.json         Vercel 函数配置
```

## 注意

- 当前版本按深信服 HCI 巡检报告结构优化。
- 如果后续巡检模板字段变化较大，可以在 `api/convert.py` 中调整识别规则。
- 上传文件仅用于本次生成，程序不落盘保存用户报告。
