# 变动记录 / Change Log

## 项目：迪士尼乐园路径规划可视化系统

**创建时间**：2026-05-03  
**源文件**：`C:\Users\fan\disney5第二问.ipynb`  
**目标目录**：`D:\桌面\可视化界面尝试\`

---

## 一、使用的权限 / 工具

| 序号 | 工具/权限 | 用途 | 备注 |
|------|-----------|------|------|
| 1 | PowerShell | 创建目录、检查数据文件是否存在 | 纯读取+创建目录，无修改 |
| 2 | Read | 读取源notebook文件 `disney5第二问.ipynb` | 理解原始代码逻辑 |
| 3 | Write | 创建 `app.py`、`index.html`、`启动系统.ipynb`、`CHANGELOG.md` | 全部为新文件，未修改任何已有文件 |
| 4 | Glob/Grep | 未使用 | - |

**关键说明**：本次操作**未删除、修改或覆盖任何已有文件**。所有新文件均创建在独立的 `D:\桌面\可视化界面尝试\` 目录中。

---

## 二、创建的文件清单

```
D:\桌面\可视化界面尝试\
├── app.py                    # HTTP后端（纯Python标准库，无需Flask）
├── templates\
│   └── index.html            # 迪士尼主题前端界面
├── 启动系统.ipynb             # Jupyter Notebook启动器
└── CHANGELOG.md              # 本文件：变动记录
```

---

## 三、对原始代码的改动说明

### 3.1 架构变更

| 项目 | 原始 notebook | 新系统 |
|------|--------------|--------|
| 交互方式 | `input()` 命令行交互 | Web网页可视化交互 |
| 输出方式 | `print()` + `display()` | JSON API + HTML表格 |
| 运行方式 | Jupyter逐Cell运行 | Python HTTP服务+浏览器访问 |
| 数据输入 | 逐行手动输入编号 | 下拉框/多选框/时间选择器 |

### 3.2 代码逻辑：完全保留，未改动

原始notebook的三个Cell核心算法，已完整移植到 `app.py` 中：

| Cell | 功能 | 在 app.py 中的位置 | 改动说明 |
|------|------|-------------------|----------|
| Cell 1 | 数据加载+名称映射+排队数据构造 | 第1-3节 `# 1. 工具函数` ~ `# 3. 名称映射建立` | 无逻辑改动。原Cell 1输出Excel改为无输出（数据直接保留在内存） |
| Cell 2 | 九维动态满意度 Top5 计算 | 第4节 `# 4. 核心算法函数` 中的 `calc_candidate_base()`, `finalize_dynamic_scores()` | 无逻辑改动。原`input()`交互改为HTTP POST参数 |
| Cell 3 | 三项目路线规划模拟 | 第4节中的 `simulate_route_sequence()`, 及 `/api/calculate` 路由中的Phase 2部分 | 无逻辑改动。保留了降级策略、21:30截止约束等全部规则 |

### 3.3 关键算法保留确认

以下核心算法与原始notebook完全一致：

- ✅ 项目名称标准化 `normalize_name()`
- ✅ 中英文名称映射（人工映射表 + 模糊匹配）
- ✅ 八维属性名称对齐（别名映射 + 包含匹配 + 模糊匹配）
- ✅ 园区图最短路算法（Dijkstra）
- ✅ 距离计算规则（同园区60m、跨园区最短路、入口60m+米奇大街）
- ✅ 步行时间 = 距离 / 速度 × 拥挤系数
- ✅ 实时排队向下取整规则（9:00-9:59使用9:00数据）
- ✅ 三类游客理想点 mu_dict 和权重 K_dict（完全一致）
- ✅ 前八维：核心六维算术平均 + z_food + z_shop
- ✅ 第九维：x9 = exp[-max(0, r-1)]，r = 实时总时间/历史总时间
- ✅ 半衰减参数 eta=2.0，K9 = ln(2)/(1-exp(-(eta-1)))^2
- ✅ 动态满意度：U = exp[-(八维指数项 + K9*(x9-1)^2)]
- ✅ 路线降级策略：保留最长可执行前缀
- ✅ 约束：21:30前到达排队、购物进入早于21:30、亲子身高限制
- ✅ 路线排序：优先完成项目数多，同等则总耗时少

### 3.4 数据来源：完全不变

所有数据仍读取原始Excel文件，路径未变：

```
C:\Users\fan\Disney4\disney_2026_05_01_hourly_wait_clean.xlsx
C:\Users\fan\Disney2（3）\修正餐饮购物0.6后的8维向量数据_20260502_231148.xlsx
C:\Users\fan\Disney3（2）\上海迪士尼_QueueTimes_20240301_20260301.xlsx
C:\Users\fan\Disney3（2）\最终项目数据表_中英文对应_填充版3.2.xlsx
```

---

## 四、新增功能

1. **可视化Web界面**：迪士尼主题（城堡、星空、金色魔法色调）
2. **时间选择器**：用 `<input type="time">` 替代手动输入
3. **游客类型下拉框**：亲子/情侣/普通，含图标
4. **位置搜索选择器**：所有项目+入口+米奇大街
5. **多选已完成项目**：按住Ctrl多选，直观方便
6. **数据总览面板**：只读显示当天开放项目、日期类型等
7. **Top5排名表**：带金银铜牌样式的排名标识
8. **路线卡片视图**：每条路线一张卡片，点击查看时间线
9. **时间线视图**：直观展示步行→排队→游玩→就餐→购物全流程
10. **剔除原因展示**：清楚说明哪些项目为何被排除
11. **响应式设计**：适配不同屏幕尺寸

---

## 五、启动方式

### 方式一：Jupyter Notebook
打开 `启动系统.ipynb`，运行Cell 1即可

### 方式二：Spyder
打开 `app.py`，按F5运行

### 方式三：命令行
```bash
cd "D:\桌面\可视化界面尝试"
python app.py
```

然后在浏览器中访问 `http://127.0.0.1:5000`

---

## 六、依赖项

依赖项（均为 Jupyter/Spyder 标准环境自带，无需额外安装）：
- pandas
- numpy
- openpyxl
- Python 标准库（http.server, json, math, re, heapq, itertools, difflib 等）

**技术选型说明**：后端采用 Python 标准库 `http.server` 模块实现 HTTP API 服务，**不依赖 Flask 或任何第三方 Web 框架**，确保在任何能运行 pandas 的 Python 环境中开箱即用。

---

## 七、技术实现细节

### 7.1 后端架构（app.py）
- **Web框架**：Python标准库 `http.server.BaseHTTPRequestHandler`（非Flask）
- **路由处理**：自定义 `DisneyRequestHandler` 类，通过 `do_GET` / `do_POST` 分发
- **JSON序列化**：使用 `json.dumps(default=str)` 处理 numpy 类型
- **CORS支持**：手动设置 `Access-Control-Allow-Origin: *` 响应头
- **HTML模板**：启动时从 `templates/index.html` 读取到内存

### 7.2 前端架构（index.html）
- **主题**：迪士尼魔法王国（深蓝紫渐变、金色点缀、星空动画）
- **核心技术**：原生 HTML/CSS/JS，无任何前端框架依赖
- **Google Fonts**：Noto Sans SC（中文字体）
- **API通信**：`fetch()` 调用后端 REST API
- **响应式**：CSS Grid 双栏布局，适配移动端

### 7.3 数据流
```
用户浏览器 ←→ HTTP :5000 ←→ app.py (服务启动时加载Excel → 内存)
   fetch()
index.html (纯前端渲染，无服务器端模板)
```
