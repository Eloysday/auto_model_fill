FROM docker.1ms.run/library/python:3.12-slim

# Python环境优化参数
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# 完整替换为阿里云Debian源，清空旧源文件避免残留国外地址
RUN echo 'deb https://mirrors.aliyun.com/debian trixie main contrib non-free' > /etc/apt/sources.list && \
    echo 'deb https://mirrors.aliyun.com/debian trixie-updates main contrib non-free' >> /etc/apt/sources.list && \
    echo 'deb https://mirrors.aliyun.com/debian-security trixie-security main contrib non-free' >> /etc/apt/sources.list && \
    rm -f /etc/apt/sources.list.d/* && \
    apt-get update -qq && \
    apt-get install -y --no-install-recommends -qq fonts-noto-cjk curl && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 分层缓存依赖，阿里云pip源加速
COPY requirements.txt .
RUN pip install --no-cache-dir \
    -i https://mirrors.aliyun.com/pypi/simple/ \
    --trusted-host mirrors.aliyun.com \
    -r requirements.txt

# 复制项目全部代码
COPY . .

# 创建持久化目录
RUN mkdir -p jobs charts

EXPOSE 8000

# 生产默认启动入口（compose会覆盖为uvicorn热重载）
CMD ["python", "server.py"]
