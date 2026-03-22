# 流水分析系统

基于多 Agent 架构的智能金融文件分析工具。上传客户 ZIP 压缩包，自动解析银行流水、财务报表、征信报告等，生成 Word 分析报告。

## 功能

- 🏦 **银行流水**：支持 XLS / XLSX / CSV / PDF（含农商行）
- 📊 **财务报表**：资产负债表、利润表（PDF + XLS）
- 📋 **征信报告**：企业/个人征信 PDF
- 🏠 **房产证**：权利人、坐落、面积（支持 LLM Vision 识别扫描件）
- 🧾 **完税证明**：税种、期间、金额
- 🤖 **AI 分析**：基于解析数据生成授信分析报告

---

## 快速开始（本地开发）

```bash
git clone https://github.com/Libaizaima/-skill.git
cd -skill

# 创建虚拟环境并安装依赖
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# 创建配置文件（见下方说明）
cp config.json.example config.json      # 填入 API key
cp web_config.json.example web_config.json  # 填入账号密码

# 启动 Web 服务（开发模式）
.venv/bin/python server.py
# 访问 http://localhost:7963
```

---

## 配置文件

> ⚠️ 这两个文件已加入 `.gitignore`，需在每台服务器上手动创建。

### `config.json` — LLM 配置

```json
{
  "api_key": "sk-xxx",
  "base_url": "https://你的中转地址/v1",
  "model": "gpt-4o",
  "temperature": 0.1,
  "max_tokens": 4096
}
```

### `web_config.json` — 网站账号

```json
{
  "username": "admin",
  "password": "改成强密码",
  "secret_key": "随机字符串至少32位"
}
```

---

## 服务器部署

### 1. 拉取代码 & 安装依赖

```bash
cd /opt
sudo git clone https://github.com/Libaizaima/-skill.git 流水分析
cd 流水分析

python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

### 2. 创建配置文件

```bash
# 手动创建 config.json 和 web_config.json（参考上方格式）
nano config.json
nano web_config.json
```

### 3. 设置目录权限

```bash
sudo chown -R www-data:www-data /opt/流水分析
sudo mkdir -p /opt/流水分析/input /opt/流水分析/output /opt/流水分析/web/db
sudo chown -R www-data:www-data /opt/流水分析/input /opt/流水分析/output /opt/流水分析/web/db
```

### 4. 注册 systemd 服务

```bash
sudo nano /etc/systemd/system/analysis-web.service
```

```ini
[Unit]
Description=流水分析系统 Web 服务
After=network.target

[Service]
User=www-data
WorkingDirectory=/opt/流水分析
ExecStart=/opt/流水分析/.venv/bin/gunicorn -w 2 -b 127.0.0.1:7963 --timeout 300 server:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable analysis-web
sudo systemctl start analysis-web
```

### 5. Nginx 反向代理

```nginx
server {
    listen 80;
    server_name 你的域名或IP;

    client_max_body_size 500M;   # 允许大文件上传

    location / {
        proxy_pass http://127.0.0.1:7963;
        proxy_read_timeout 600s;
        proxy_send_timeout 600s;
        proxy_buffering off;     # SSE 实时日志必须关闭缓冲
        proxy_cache off;
    }
}
```

```bash
sudo nginx -t && sudo systemctl reload nginx
```

> 如通过 **frp 反代**访问，需同时在 VPS 侧 nginx 配置 `client_max_body_size 500M`。

---

## 日常运维

| 操作 | 命令 |
|------|------|
| 更新代码 | `cd /opt/流水分析 && git pull` |
| 重启服务 | `sudo systemctl restart analysis-web` |
| 查看实时日志 | `sudo journalctl -u analysis-web -f` |
| 查看错误 | `sudo journalctl -u analysis-web -n 50 --no-pager` |

> 只有 `requirements.txt` 有变化时才需重新 `pip install`，平时 `git pull` + `restart` 即可。

---

## 常见问题

| 现象 | 原因 | 解决 |
|------|------|------|
| 上传一直显示"上传中" | Nginx `client_max_body_size` 太小（默认 1MB）| 改为 `500M` |
| 500 Internal Server Error | `input/` 目录无写权限 | `chown -R www-data /opt/流水分析` |
| Worker Timeout | gunicorn 默认超时 30s，大文件上传太慢 | 加 `--timeout 300` 参数 |
| 文件下载"暂无可下载文件" | 旧记录无 output_dir，自动按公司名匹配 | 正常情况，刷新页面重试 |
| LLM 调用失败 | API Key 错误或网络不通 | 检查 `config.json`，LLM 不可用时降级为模板 |

---

## 目录说明

```
├── server.py          # Web 服务入口（Flask）
├── src/               # 核心分析引擎
│   ├── main.py        # 命令行入口
│   ├── agents/        # Brain Agent + Tool Agent
│   ├── *_parser.py    # 各类文件解析器
│   ├── analyzer.py    # 统计分析
│   ├── ai_analyzer.py # AI 智能分析
│   └── report_generator.py  # DOCX 报告生成
├── web/templates/     # 前端页面
├── requirements.txt   # Python 依赖
├── MAINTENANCE.md     # 详细开发维护文档
└── .gitignore
```

详细的解析器开发文档、架构说明请参阅 [MAINTENANCE.md](MAINTENANCE.md)。
