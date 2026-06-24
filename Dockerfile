FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK=True \
    PORT=7860

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements-server.txt ./
RUN pip install --no-cache-dir -r requirements-server.txt

# 构建镜像时预下载 PP-OCRv6 模型，避免首次使用等待。
RUN python -c "from paddleocr import PaddleOCR; PaddleOCR(use_doc_orientation_classify=False, use_doc_unwarping=False, use_textline_orientation=False, engine='onnxruntime', engine_config={'providers':['CPUExecutionProvider']})"

COPY . .

EXPOSE 7860
CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${PORT}"]
