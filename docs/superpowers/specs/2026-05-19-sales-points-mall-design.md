# 销售积分排名及商城系统 — 设计规格

## 概述

解析「东4月.xlsx」等月度销售数据文件，按每 100 元销售金额 = 1 积分（四舍五入）计算积分，提供排名统计、一键导出、积分商城兑换功能。

- **管理员**：何莎山、吴泽光（硬编码）
- **技术栈**：Python Flask + SQLite + Bootstrap 5 + Jinja2 模板
- **部署**：免费平台（Render / Railway）

---

## 数据库设计

6 张表：

### users
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| name | TEXT UNIQUE | 姓名，唯一 |
| is_admin | INTEGER | 0/1，管理员由代码判断 |
| points | INTEGER | 积分余额 |
| created_at | TIMESTAMP | 注册时间 |

### sales_records
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| date | TEXT | 日期 |
| biz_org_name | TEXT | 业务机构名称 |
| product_name | TEXT | 商品名称 |
| salesperson | TEXT | 营业员名字 |
| quantity | REAL | 销售数量 |
| amount | REAL | 销售金额（列 T） |
| points | INTEGER | 该条记录对应积分 = round(amount/100) |
| row_hash | TEXT UNIQUE | 去重哈希 |
| import_id | INTEGER FK | 关联 import_logs |

### products
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| name | TEXT | 商品名称 |
| description | TEXT | 描述 |
| image_url | TEXT | 图片链接 |
| points_required | INTEGER | 所需积分 |
| stock | INTEGER | 库存（-1 为无限） |
| is_active | INTEGER | 1=上架 0=下架 |
| created_at | TIMESTAMP | 创建时间 |

### orders
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| user_id | INTEGER FK | 用户 ID |
| product_id | INTEGER FK | 商品 ID |
| quantity | INTEGER | 兑换数量 |
| points_spent | INTEGER | 消耗积分 |
| contact_name | TEXT | 联系人姓名 |
| contact_phone | TEXT | 联系电话 |
| address | TEXT | 收货地址 |
| status | TEXT | 待处理/处理中/已发货/已完成 |
| created_at | TIMESTAMP | 下单时间 |

### points_log
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| user_id | INTEGER FK | 用户 ID |
| source | TEXT | 来源（销售导入/兑换消费/管理员调整） |
| points_change | INTEGER | 积分变动（正=增加，负=扣减） |
| balance_after | INTEGER | 变动后余额 |
| created_at | TIMESTAMP | 时间 |

### import_logs
| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增 |
| filename | TEXT | 文件名 |
| record_count | INTEGER | 导入记录数 |
| total_amount | REAL | 导入总金额 |
| created_at | TIMESTAMP | 导入时间 |

---

## 页面与路由

| 路由 | 页面 | 权限 |
|------|------|------|
| `/` | 首页仪表盘（总销售额、总积分、Top10 排名） | 所有人 |
| `/login` | 登录（输入姓名，无密码） | 未登录 |
| `/register` | 注册（输入姓名） | 未登录 |
| `/import` | 数据导入（上传 xlsx → 预览 → 确认导入） | 仅管理员 |
| `/ranking` | 排名（按营业员/机构/商品筛选，支持导出 Excel） | 所有人 |
| `/mall` | 积分商城（商品列表，普通用户可兑换） | 已登录 |
| `/my-points` | 我的积分（余额 + 流水） | 已登录 |
| `/my-orders` | 我的订单 | 已登录 |
| `/admin` | 管理后台（商品管理 + 订单管理 + 用户管理） | 仅管理员 |
| `/admin/products` | 商品 CRUD | 仅管理员 |
| `/admin/orders` | 所有订单管理 | 仅管理员 |

---

## 核心逻辑

### 积分计算
- `points = round(销售金额 / 100)`，四舍五入
- 销售金额取自 Excel 列 T（第 19 列）
- 积分在导入时一次性计算并写入 sales_records

### Excel 导入去重
- Hash = MD5(日期 + 业务机构名称 + 商品名称 + 营业员 + 销售数量 + 销售金额)
- 同一 hash 跳过，支持多次导入同一/不同月度文件
- 导入记录写入 import_logs
- 对应营业员累计积分自动更新 users.points

### 管理员判断
```python
ADMIN_NAMES = ['何莎山', '吴泽光']
# 登录时检查 name 是否在此列表中，设置 session['is_admin']
```

### 排名与导出
- 按营业员 / 业务机构 / 商品三种维度汇总
- 支持月份筛选（根据 date 字段 LIKE '2026/04%' 等）
- 导出：Flask 生成 xlsx 文件，提供下载链接

### 兑换流程
1. 用户浏览 `/mall`，点击商品
2. 填写兑换数量、联系人、电话、地址
3. 系统校验积分余额，扣除积分，生成订单（状态：待处理）
4. 记录 points_log（source=兑换消费，负值）
5. 管理员在 `/admin/orders` 查看并处理

---

## 技术细节

### 目录结构
```
project/
├── app.py              # Flask 主应用
├── models.py           # SQLAlchemy 模型
├── utils.py            # Excel 解析、积分计算工具
├── requirements.txt    # Flask, openpyxl, SQLAlchemy, etc.
├── templates/          # Jinja2 模板
│   ├── base.html       # 公共布局（导航栏）
│   ├── index.html      # 仪表盘
│   ├── login.html      # 登录
│   ├── register.html   # 注册
│   ├── import.html     # 数据导入
│   ├── ranking.html    # 排名
│   ├── mall.html       # 商城
│   ├── my_points.html  # 我的积分
│   ├── my_orders.html  # 我的订单
│   └── admin/          # 管理后台模板
│       ├── products.html
│       └── orders.html
└── static/
    └── (Bootstrap 5 CDN，无需本地文件)
```

### Session 管理
- Flask session（服务端 cookie）
- 无需密码，name 登录直接设置 session['user_id'] 和 session['is_admin']
- 未登录用户重定向到 `/login`

### 安全要点
- 管理员路由使用装饰器检查 `session['is_admin']`
- 订单查找、积分查询根据 `session['user_id']` 限制当前用户范围
- Excel 上传文件大小限制 10MB

---

## 部署

- 平台：Render（render.com）或 Railway（railway.app）
- Python 3.10+
- 启动命令：`gunicorn app:app`
- SQLite 数据库文件放在项目目录下
- 无需外部数据库服务
