---
title: PDF Smart Splitter
emoji: 📄
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 7860
pinned: false
---

# PDF 智能拆分与重命名

网页在浏览器中完成 PDF 分界扫描、拆分、重命名和 Excel 目录生成，仅将分界页图片发送到同站的 PP-OCRv6 服务识别。

## 本地启动

```bash
docker build -t pdf-smart-splitter .
docker run --rm -p 7860:7860 pdf-smart-splitter
```

打开 `http://127.0.0.1:7860`。

## 健康检查

```bash
curl http://127.0.0.1:7860/api/health
```
