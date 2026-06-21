FROM python:3.12-slim

# 替换Debian中科大国内源，解决apt更新安装卡顿
RUN echo "deb http://mirrors.ustc.edu.cn/debian trixie main" > /etc/apt/sources.list \
    && echo "deb http://mirrors.ustc.edu.cn/debian trixie-updates main" >> /etc/apt/sources.list \
    && echo "deb http://mirrors.ustc.edu.cn/debian-security trixie-security main" >> /etc/apt/sources.list

WORKDIR /app

# 安装系统依赖并清理缓存，减小镜像体积
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk curl \
    && rm -rf /var/lib/apt/lists/*

# 分层缓存：先复制依赖，代码修改不会重复重装pip包
COPY requirements.txt .
# pip使用清华源加速下载依赖
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 复制项目代码
COPY . .

# 创建持久化目录
RUN mkdir -p jobs charts

EXPOSE 8000

# 生产默认启动命令（compose会覆盖为uvicorn热重载命令，不冲突）
CMD ["python", "server.py"]
