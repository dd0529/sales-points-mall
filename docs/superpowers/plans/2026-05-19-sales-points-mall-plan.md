# 销售积分排名及商城系统 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个完整的 Flask 积分排名及商城系统，支持 Excel 导入、积分计算、排名导出、积分商城兑换。

**Architecture:** Flask 单体应用，SQLite 数据库，Bootstrap 5 + Jinja2 模板渲染。Session 认证，装饰器控制权限。openpyxl 解析 Excel。无前端构建步骤。

**Tech Stack:** Python 3.10+, Flask 3.x, Flask-SQLAlchemy 3.x, openpyxl, Bootstrap 5 (CDN), SQLite

---

### Task 1: 项目骨架

**Files:**
- Create: `requirements.txt`
- Create: `app.py`
- Create: `Procfile`

- [ ] **Step 1: 创建 requirements.txt**

```
flask==3.1.0
flask-sqlalchemy==3.1.1
openpyxl==3.1.5
gunicorn==23.0.0
```

- [ ] **Step 2: 创建 Procfile**

```
web: gunicorn app:app
```

- [ ] **Step 3: 创建 app.py 骨架**

```python
import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(basedir, "data.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

db = SQLAlchemy(app)

@app.route('/')
def index():
    return 'Hello'

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
```

- [ ] **Step 4: 安装依赖并验证启动**

```bash
pip install -r requirements.txt
python app.py
```

Expected: `Running on http://127.0.0.1:5000`

---

### Task 2: 数据库模型

**Files:**
- Create: `models.py`
- Modify: `app.py`

- [ ] **Step 1: 创建 models.py**

```python
from app import db
from datetime import datetime

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    is_admin = db.Column(db.Integer, default=0)
    points = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class SalesRecord(db.Model):
    __tablename__ = 'sales_records'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(20))
    biz_org_name = db.Column(db.String(200))
    product_name = db.Column(db.String(100))
    salesperson = db.Column(db.String(50))
    quantity = db.Column(db.Float, default=0)
    amount = db.Column(db.Float, default=0)
    points = db.Column(db.Integer, default=0)
    row_hash = db.Column(db.String(64), unique=True)
    import_id = db.Column(db.Integer, db.ForeignKey('import_logs.id'))

class Product(db.Model):
    __tablename__ = 'products'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    image_url = db.Column(db.String(500))
    points_required = db.Column(db.Integer, nullable=False)
    stock = db.Column(db.Integer, default=-1)
    is_active = db.Column(db.Integer, default=1)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Order(db.Model):
    __tablename__ = 'orders'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'))
    quantity = db.Column(db.Integer, default=1)
    points_spent = db.Column(db.Integer, default=0)
    contact_name = db.Column(db.String(50))
    contact_phone = db.Column(db.String(30))
    address = db.Column(db.String(300))
    status = db.Column(db.String(20), default='待处理')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    user = db.relationship('User', backref='orders')
    product = db.relationship('Product', backref='orders')

class PointsLog(db.Model):
    __tablename__ = 'points_log'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    source = db.Column(db.String(50))
    points_change = db.Column(db.Integer, default=0)
    balance_after = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    user = db.relationship('User', backref='points_logs')

class ImportLog(db.Model):
    __tablename__ = 'import_logs'
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200))
    record_count = db.Column(db.Integer, default=0)
    total_amount = db.Column(db.Float, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    records = db.relationship('SalesRecord', backref='import_log')
```

- [ ] **Step 2: 在 app.py 中导入模型（替换原有代码）**

```python
import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(basedir, "data.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

db = SQLAlchemy(app)

# models must be imported after db init
from models import User, SalesRecord, Product, Order, PointsLog, ImportLog

@app.route('/')
def index():
    return 'Hello'

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
```

- [ ] **Step 3: 验证数据库创建**

```bash
python -c "from app import app, db; app.app_context().push(); db.create_all(); print('DB created')"
```

Expected: `DB created`，且项目目录生成 `data.db`

---

### Task 3: 认证系统

**Files:**
- Create: `auth.py`
- Modify: `app.py`
- Create: `templates/login.html`
- Create: `templates/register.html`

- [ ] **Step 1: 创建 auth.py**

```python
from functools import wraps
from flask import session, redirect, url_for

ADMIN_NAMES = ['何莎山', '吴泽光']

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if not session.get('is_admin'):
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated
```

- [ ] **Step 2: 在 app.py 添加路由**

```python
# 添加在 imports 后
from auth import login_required, admin_required, ADMIN_NAMES
from models import User, SalesRecord, Product, Order, PointsLog, ImportLog

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            return render_template('login.html', error='请输入姓名')
        user = User.query.filter_by(name=name).first()
        if not user:
            return render_template('login.html', error='用户不存在，请先注册')
        session['user_id'] = user.id
        session['user_name'] = user.name
        session['is_admin'] = name in ADMIN_NAMES
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            return render_template('register.html', error='请输入姓名')
        if User.query.filter_by(name=name).first():
            return render_template('register.html', error='该姓名已注册')
        user = User(name=name, is_admin=1 if name in ADMIN_NAMES else 0)
        db.session.add(user)
        db.session.commit()
        session['user_id'] = user.id
        session['user_name'] = user.name
        session['is_admin'] = name in ADMIN_NAMES
        return redirect(url_for('index'))
    return render_template('register.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))
```

注意：需要添加 `request` 和 `render_template` 的 import：
```python
from flask import Flask, render_template, request, redirect, url_for, session, send_file
```

- [ ] **Step 3: 创建 templates/login.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>登录 - 积分系统</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container" style="max-width:400px;margin-top:100px;">
    <div class="card shadow">
        <div class="card-body p-4">
            <h3 class="text-center mb-4">积分排名系统</h3>
            {% if error %}
            <div class="alert alert-danger">{{ error }}</div>
            {% endif %}
            <form method="POST">
                <div class="mb-3">
                    <label class="form-label">姓名</label>
                    <input type="text" name="name" class="form-control" placeholder="请输入您的姓名" required>
                </div>
                <button type="submit" class="btn btn-primary w-100">登录</button>
            </form>
            <div class="text-center mt-3">
                <a href="{{ url_for('register') }}">还没有账号？立即注册</a>
            </div>
        </div>
    </div>
</div>
</body>
</html>
```

- [ ] **Step 4: 创建 templates/register.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>注册 - 积分系统</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container" style="max-width:400px;margin-top:100px;">
    <div class="card shadow">
        <div class="card-body p-4">
            <h3 class="text-center mb-4">用户注册</h3>
            {% if error %}
            <div class="alert alert-danger">{{ error }}</div>
            {% endif %}
            <form method="POST">
                <div class="mb-3">
                    <label class="form-label">姓名</label>
                    <input type="text" name="name" class="form-control" placeholder="请输入您的真实姓名" required>
                </div>
                <button type="submit" class="btn btn-success w-100">注册</button>
            </form>
            <div class="text-center mt-3">
                <a href="{{ url_for('login') }}">已有账号？去登录</a>
            </div>
        </div>
    </div>
</div>
</body>
</html>
```

- [ ] **Step 5: 验证**

启动 `python app.py`，访问 `http://127.0.0.1:5000/register`，注册一个测试用户，确认跳转到首页。

---

### Task 4: 公共布局模板 + 导航栏

**Files:**
- Create: `templates/base.html`
- Create: `static/style.css`

- [ ] **Step 1: 创建 templates/base.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}积分系统{% endblock %}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="{{ url_for('static', filename='style.css') }}" rel="stylesheet">
</head>
<body>
<nav class="navbar navbar-expand-lg navbar-dark bg-primary">
    <div class="container">
        <a class="navbar-brand" href="{{ url_for('index') }}">积分系统</a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNav">
            <ul class="navbar-nav me-auto">
                {% if session.user_id %}
                <li class="nav-item"><a class="nav-link" href="{{ url_for('index') }}">首页</a></li>
                <li class="nav-item"><a class="nav-link" href="{{ url_for('ranking') }}">排名</a></li>
                <li class="nav-item"><a class="nav-link" href="{{ url_for('mall') }}">积分商城</a></li>
                <li class="nav-item"><a class="nav-link" href="{{ url_for('my_points') }}">我的积分</a></li>
                <li class="nav-item"><a class="nav-link" href="{{ url_for('my_orders') }}">我的订单</a></li>
                {% if session.is_admin %}
                <li class="nav-item dropdown">
                    <a class="nav-link dropdown-toggle" href="#" data-bs-toggle="dropdown">管理</a>
                    <ul class="dropdown-menu">
                        <li><a class="dropdown-item" href="{{ url_for('import_data') }}">数据导入</a></li>
                        <li><a class="dropdown-item" href="{{ url_for('admin_products') }}">商品管理</a></li>
                        <li><a class="dropdown-item" href="{{ url_for('admin_orders') }}">订单管理</a></li>
                    </ul>
                </li>
                {% endif %}
                {% endif %}
            </ul>
            <ul class="navbar-nav">
                {% if session.user_id %}
                <li class="nav-item"><span class="nav-link text-light">{{ session.user_name }}{% if session.is_admin %} (管理员){% endif %}</span></li>
                <li class="nav-item"><a class="nav-link" href="{{ url_for('logout') }}">退出</a></li>
                {% else %}
                <li class="nav-item"><a class="nav-link" href="{{ url_for('login') }}">登录</a></li>
                {% endif %}
            </ul>
        </div>
    </div>
</nav>
<div class="container mt-4">
    {% block content %}{% endblock %}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
```

- [ ] **Step 2: 创建 static/style.css**

```css
.card-stats {
    border-left: 4px solid #0d6efd;
    transition: transform 0.2s;
}
.card-stats:hover { transform: translateY(-2px); }
.ranking-table th { position: sticky; top: 0; background: #fff; z-index: 1; }
```

- [ ] **Step 3: 更新 login.html 和 register.html 不继承 base.html（独立页面，无需导航栏），无需修改。**

---

### Task 5: 首页仪表盘

**Files:**
- Create: `templates/index.html`
- Modify: `app.py` — 更新 index 路由

- [ ] **Step 1: 更新 app.py 中的 index 路由**

```python
from sqlalchemy import func

@app.route('/')
def index():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    total_sales = db.session.query(func.sum(SalesRecord.amount)).scalar() or 0
    total_points = db.session.query(func.sum(SalesRecord.points)).scalar() or 0
    total_records = SalesRecord.query.count()
    
    # Top 10 salespersons by amount
    top_sales = db.session.query(
        SalesRecord.salesperson,
        func.sum(SalesRecord.amount).label('total_amt'),
        func.sum(SalesRecord.quantity).label('total_qty'),
        func.sum(SalesRecord.points).label('total_pts')
    ).group_by(SalesRecord.salesperson).order_by(
        func.sum(SalesRecord.amount).desc()
    ).limit(10).all()
    
    return render_template('index.html',
        total_sales=total_sales, total_points=total_points,
        total_records=total_records, top_sales=top_sales)
```

- [ ] **Step 2: 创建 templates/index.html**

```html
{% extends "base.html" %}
{% block title %}首页 - 积分系统{% endblock %}
{% block content %}
<h2 class="mb-4">仪表盘</h2>
<div class="row g-3 mb-4">
    <div class="col-md-4">
        <div class="card card-stats shadow-sm">
            <div class="card-body">
                <h6 class="text-muted">总销售金额</h6>
                <h3 class="text-primary">¥{{ "%.2f"|format(total_sales) }}</h3>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-stats shadow-sm">
            <div class="card-body">
                <h6 class="text-muted">总积分</h6>
                <h3 class="text-success">{{ total_points }}</h3>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-stats shadow-sm">
            <div class="card-body">
                <h6 class="text-muted">销售记录数</h6>
                <h3 class="text-warning">{{ total_records }}</h3>
            </div>
        </div>
    </div>
</div>

<div class="card shadow-sm">
    <div class="card-header"><h5 class="mb-0">销售金额 Top 10</h5></div>
    <div class="card-body p-0">
        <table class="table table-striped mb-0">
            <thead class="ranking-table"><tr><th>排名</th><th>营业员</th><th>销售数量</th><th>销售金额</th><th>积分</th></tr></thead>
            <tbody>
                {% for i, s in enumerate(top_sales, 1) %}
                <tr>
                    <td><span class="badge bg-{{ 'warning' if i<=3 else 'secondary' }}">{{ i }}</span></td>
                    <td>{{ s.salesperson }}</td>
                    <td>{{ s.total_qty|int }}</td>
                    <td>¥{{ "%.2f"|format(s.total_amt) }}</td>
                    <td>{{ s.total_pts }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
{% endblock %}
```

- [ ] **Step 1.5: 在 app.py 中添加 Jinja2 上下文处理器（添加在路由之前）**

```python
@app.context_processor
def utility_processor():
    return dict(enumerate=enumerate)
```

---

### Task 6: Excel 解析工具 + 导入后端

**Files:**
- Create: `utils.py`
- Modify: `app.py` — 添加 `/import` 路由

- [ ] **Step 1: 创建 utils.py**

```python
import hashlib
import openpyxl

def parse_excel(filepath):
    """Parse Excel file, return list of dicts and any error message."""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
    except Exception as e:
        return None, f'文件无法打开: {str(e)}'

    # Find header row (Row 1 per the sample file)
    header_row = [str(c.value) if c.value is not None else '' for c in ws[1]]

    # Identify key columns by name substring matching
    col_map = {}
    for key, labels in [
        ('date', ['日期']),
        ('biz_org_name', ['业务机构名称']),
        ('product_name', ['商品名称']),
        ('salesperson', ['营业员名字']),
        ('quantity', ['销售数量']),
        ('amount', ['销售金额']),
    ]:
        for i, h in enumerate(header_row):
            for label in labels:
                if label in h:
                    # Prevent matching '不含税销售金额1' when looking for '销售金额'
                    if key == 'amount' and '不含税' in h:
                        continue
                    col_map[key] = i
                    break
            if key in col_map:
                break

    required = ['biz_org_name', 'product_name', 'salesperson', 'quantity', 'amount']
    missing = [k for k in required if k not in col_map]
    if missing:
        return None, f'缺少必要列: {", ".join(missing)}'

    records = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        try:
            amount = float(row[col_map['amount']]) if row[col_map['amount']] else 0
        except (ValueError, TypeError):
            amount = 0
        try:
            qty = float(row[col_map['quantity']]) if row[col_map['quantity']] else 0
        except (ValueError, TypeError):
            qty = 0

        if amount <= 0:
            continue

        date_val = str(row[col_map['date']]) if 'date' in col_map and row[col_map['date']] else ''
        biz = str(row[col_map['biz_org_name']]) if row[col_map['biz_org_name']] else ''
        prod = str(row[col_map['product_name']]) if row[col_map['product_name']] else ''
        person = str(row[col_map['salesperson']]) if row[col_map['salesperson']] else ''

        hash_str = f'{date_val}|{biz}|{prod}|{person}|{qty}|{amount}'
        row_hash = hashlib.md5(hash_str.encode()).hexdigest()

        records.append({
            'date': date_val,
            'biz_org_name': biz,
            'product_name': prod,
            'salesperson': person,
            'quantity': qty,
            'amount': amount,
            'points': round(amount / 100),
            'row_hash': row_hash,
        })

    return records, None
```

- [ ] **Step 2: 在 app.py 添加导入路由**

```python
import os
from werkzeug.utils import secure_filename
from utils import parse_excel

@app.route('/import', methods=['GET', 'POST'])
@admin_required
def import_data():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename.endswith('.xlsx'):
            return render_template('import.html', error='请上传 .xlsx 文件')
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(basedir, 'uploads', filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        file.save(filepath)

        records, error = parse_excel(filepath)
        os.remove(filepath)

        if error:
            return render_template('import.html', error=error)
        if not records:
            return render_template('import.html', error='文件中没有有效数据')

        # Preview mode: store in session for confirmation
        session['import_preview'] = records
        return render_template('import.html', preview=True, 
            records=records[:20], total_rows=len(records),
            total_amount=sum(r['amount'] for r in records),
            total_points=sum(r['points'] for r in records))

    return render_template('import.html', preview=False)


@app.route('/import/confirm', methods=['POST'])
@admin_required
def import_confirm():
    records = session.pop('import_preview', None)
    if not records:
        return redirect(url_for('import_data'))

    # Create import log first
    log = ImportLog(
        filename='import_session',
        record_count=0,
        total_amount=0
    )
    db.session.add(log)
    db.session.flush()  # get log.id

    imported = 0
    total_amount = 0
    for r in records:
        existing = SalesRecord.query.filter_by(row_hash=r['row_hash']).first()
        if existing:
            continue
        r['import_id'] = log.id
        sr = SalesRecord(**r)
        db.session.add(sr)
        imported += 1
        total_amount += r['amount']

        # Update user points
        user = User.query.filter_by(name=r['salesperson']).first()
        if user:
            user.points += r['points']
            db.session.add(PointsLog(
                user_id=user.id, source='销售导入',
                points_change=r['points'], balance_after=user.points
            ))

    log.record_count = imported
    log.total_amount = total_amount
    db.session.commit()

    return redirect(url_for('index'))
```

---

### Task 7: 数据导入页面（前端）

**Files:**
- Create: `templates/import.html`

- [ ] **Step 1: 创建 templates/import.html**

```html
{% extends "base.html" %}
{% block title %}数据导入 - 积分系统{% endblock %}
{% block content %}
<h2 class="mb-4">数据导入</h2>

{% if error %}
<div class="alert alert-danger">{{ error }}</div>
{% endif %}

{% if not preview %}
<div class="card shadow-sm">
    <div class="card-body">
        <h5>上传 Excel 文件</h5>
        <p class="text-muted">支持「东4月」等月度销售数据表格（.xlsx）</p>
        <form method="POST" enctype="multipart/form-data">
            <div class="mb-3">
                <input type="file" name="file" class="form-control" accept=".xlsx" required>
            </div>
            <button type="submit" class="btn btn-primary">上传预览</button>
        </form>
    </div>
</div>
{% else %}
<div class="card shadow-sm mb-3">
    <div class="card-body">
        <h5>预览数据</h5>
        <p>共 <strong>{{ total_rows }}</strong> 条记录，
           总金额 <strong>¥{{ "%.2f"|format(total_amount) }}</strong>，
           总积分 <strong>{{ total_points }}</strong></p>
        <form method="POST" action="{{ url_for('import_confirm') }}">
            <button type="submit" class="btn btn-success btn-lg">确认导入</button>
            <a href="{{ url_for('import_data') }}" class="btn btn-outline-secondary">取消</a>
        </form>
    </div>
</div>

<div class="card shadow-sm">
    <div class="card-header">数据预览（前 20 条）</div>
    <div class="card-body p-0 table-responsive">
        <table class="table table-sm mb-0">
            <thead><tr><th>日期</th><th>业务机构</th><th>商品</th><th>营业员</th><th>数量</th><th>金额</th><th>积分</th></tr></thead>
            <tbody>
                {% for r in records %}
                <tr>
                    <td>{{ r.date }}</td>
                    <td>{{ r.biz_org_name }}</td>
                    <td>{{ r.product_name }}</td>
                    <td>{{ r.salesperson }}</td>
                    <td>{{ r.quantity }}</td>
                    <td>¥{{ "%.2f"|format(r.amount) }}</td>
                    <td>{{ r.points }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
{% endif %}
{% endblock %}
```

---

### Task 8: 排名页面 + 导出

**Files:**
- Create: `templates/ranking.html`
- Modify: `app.py` — 添加 `/ranking` 和 `/ranking/export` 路由

- [ ] **Step 1: 在 app.py 添加排名路由**

```python
from sqlalchemy import func, desc
from io import BytesIO

@app.route('/ranking')
def ranking():
    dim = request.args.get('dim', 'salesperson')  # salesperson / biz_org / product
    page = request.args.get('page', 1, type=int)
    per_page = 50

    dim_map = {
        'salesperson': SalesRecord.salesperson,
        'biz_org': SalesRecord.biz_org_name,
        'product': SalesRecord.product_name,
    }
    group_col = dim_map.get(dim, SalesRecord.salesperson)

    query = db.session.query(
        group_col.label('name'),
        func.sum(SalesRecord.quantity).label('total_qty'),
        func.sum(SalesRecord.amount).label('total_amt'),
        func.sum(SalesRecord.points).label('total_pts')
    ).group_by(group_col).order_by(desc('total_amt'))

    pagination = query.paginate(page=page, per_page=per_page)
    results = pagination.items

    return render_template('ranking.html',
        results=results, dim=dim, pagination=pagination,
        enumerate=enumerate)


@app.route('/ranking/export')
def ranking_export():
    dim = request.args.get('dim', 'salesperson')
    dim_map = {
        'salesperson': SalesRecord.salesperson,
        'biz_org': SalesRecord.biz_org_name,
        'product': SalesRecord.product_name,
    }
    group_col = dim_map.get(dim, SalesRecord.salesperson)

    results = db.session.query(
        group_col.label('name'),
        func.sum(SalesRecord.quantity).label('total_qty'),
        func.sum(SalesRecord.amount).label('total_amt'),
        func.sum(SalesRecord.points).label('total_pts')
    ).group_by(group_col).order_by(desc('total_amt')).all()

    import openpyxl as xl
    wb = xl.Workbook()
    ws = wb.active
    ws.title = '排名结果'
    ws.append(['排名', '名称', '销售数量', '销售金额', '积分'])
    for i, r in enumerate(results, 1):
        ws.append([i, r.name, round(r.total_qty, 2), round(r.total_amt, 2), r.total_pts])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=f'ranking_{dim}.xlsx')
```

- [ ] **Step 2: 创建 templates/ranking.html**

```html
{% extends "base.html" %}
{% block title %}排名 - 积分系统{% endblock %}
{% block content %}
<h2 class="mb-3">销售排名</h2>

<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
    <div class="btn-group">
        <a href="?dim=salesperson" class="btn btn-{{ 'primary' if dim=='salesperson' else 'outline-primary' }}">按营业员</a>
        <a href="?dim=biz_org" class="btn btn-{{ 'primary' if dim=='biz_org' else 'outline-primary' }}">按业务机构</a>
        <a href="?dim=product" class="btn btn-{{ 'primary' if dim=='product' else 'outline-primary' }}">按商品</a>
    </div>
    <a href="{{ url_for('ranking_export', dim=dim) }}" class="btn btn-success">导出 Excel</a>
</div>

<div class="card shadow-sm">
    <div class="card-body p-0 table-responsive">
        <table class="table table-striped mb-0">
            <thead class="ranking-table">
                <tr><th>排名</th><th>名称</th><th>销售数量</th><th>销售金额</th><th>积分</th></tr>
            </thead>
            <tbody>
                {% set start = (pagination.page - 1) * pagination.per_page + 1 %}
                {% for i, r in enumerate(results, start) %}
                <tr>
                    <td><span class="badge bg-{{ 'warning' if i<=3 else 'secondary' }}">{{ i }}</span></td>
                    <td>{{ r.name }}</td>
                    <td>{{ r.total_qty|round(2) }}</td>
                    <td>¥{{ "%.2f"|format(r.total_amt) }}</td>
                    <td><strong>{{ r.total_pts }}</strong></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

{% if pagination.pages > 1 %}
<nav class="mt-3">
    <ul class="pagination justify-content-center">
        {% for p in range(1, pagination.pages + 1) %}
        <li class="page-item {{ 'active' if p == pagination.page }}">
            <a class="page-link" href="?dim={{ dim }}&page={{ p }}">{{ p }}</a>
        </li>
        {% endfor %}
    </ul>
</nav>
{% endif %}
{% endblock %}
```

---

### Task 9: 积分商城 - 浏览

**Files:**
- Create: `templates/mall.html`
- Modify: `app.py` — 添加 `/mall` 路由

- [ ] **Step 1: 在 app.py 添加商城路由**

```python
@app.route('/mall')
@login_required
def mall():
    products = Product.query.filter_by(is_active=1).order_by(Product.points_required).all()
    user = User.query.get(session['user_id'])
    return render_template('mall.html', products=products, user=user)
```

- [ ] **Step 2: 创建 templates/mall.html**

```html
{% extends "base.html" %}
{% block title %}积分商城{% endblock %}
{% block content %}
<h2 class="mb-3">积分商城</h2>
<p class="text-muted">我的积分余额：<strong class="text-success fs-4">{{ user.points }}</strong></p>

<div class="row g-3">
    {% for p in products %}
    <div class="col-md-4">
        <div class="card h-100 shadow-sm">
            {% if p.image_url %}
            <img src="{{ p.image_url }}" class="card-img-top" style="height:200px;object-fit:cover;" alt="{{ p.name }}">
            {% else %}
            <div class="card-img-top bg-light d-flex align-items-center justify-content-center" style="height:200px;">
                <span class="text-muted fs-1">🎁</span>
            </div>
            {% endif %}
            <div class="card-body d-flex flex-column">
                <h5 class="card-title">{{ p.name }}</h5>
                <p class="card-text text-muted small">{{ p.description or '暂无描述' }}</p>
                <div class="mt-auto">
                    <span class="text-danger fw-bold fs-5">{{ p.points_required }} 积分</span>
                    {% if p.stock > 0 %}
                    <span class="text-muted ms-2">库存: {{ p.stock }}</span>
                    {% endif %}
                    <a href="{{ url_for('exchange', product_id=p.id) }}" class="btn btn-primary float-end">兑换</a>
                </div>
            </div>
        </div>
    </div>
    {% endfor %}
    {% if not products %}
    <div class="col-12"><div class="alert alert-info">暂无商品上架</div></div>
    {% endif %}
</div>
{% endblock %}
```

---

### Task 10: 积分商城 - 兑换

**Files:**
- Create: `templates/exchange.html`
- Modify: `app.py` — 添加兑换路由

- [ ] **Step 1: 在 app.py 添加兑换路由**

```python
@app.route('/exchange/<int:product_id>', methods=['GET', 'POST'])
@login_required
def exchange(product_id):
    product = Product.query.get_or_404(product_id)
    user = User.query.get(session['user_id'])

    if request.method == 'POST':
        qty = int(request.form.get('quantity', 1))
        if qty < 1:
            return render_template('exchange.html', product=product, user=user, error='数量无效')

        # Check stock
        if product.stock >= 0 and qty > product.stock:
            return render_template('exchange.html', product=product, user=user, error='库存不足')

        total_points = product.points_required * qty
        if user.points < total_points:
            return render_template('exchange.html', product=product, user=user, error='积分不足')

        contact_name = request.form.get('contact_name', '').strip()
        contact_phone = request.form.get('contact_phone', '').strip()
        address = request.form.get('address', '').strip()
        if not contact_name or not contact_phone or not address:
            return render_template('exchange.html', product=product, user=user, error='请填写完整的收货信息')

        # Deduct points
        user.points -= total_points
        # Deduct stock
        if product.stock >= 0:
            product.stock -= qty

        order = Order(
            user_id=user.id, product_id=product.id,
            quantity=qty, points_spent=total_points,
            contact_name=contact_name, contact_phone=contact_phone,
            address=address, status='待处理'
        )
        db.session.add(order)
        db.session.add(PointsLog(
            user_id=user.id, source='兑换消费',
            points_change=-total_points, balance_after=user.points
        ))
        db.session.commit()

        return render_template('exchange.html', product=product, user=user, success=True)

    return render_template('exchange.html', product=product, user=user)
```

- [ ] **Step 2: 创建 templates/exchange.html**

```html
{% extends "base.html" %}
{% block title %}兑换商品{% endblock %}
{% block content %}
<h2 class="mb-4">兑换商品</h2>

{% if success %}
<div class="alert alert-success">兑换成功！您的订单已提交，请等待管理员处理。</div>
<a href="{{ url_for('mall') }}" class="btn btn-primary">返回商城</a>
<a href="{{ url_for('my_orders') }}" class="btn btn-outline-primary">查看订单</a>
{% else %}
<div class="row">
    <div class="col-md-6">
        <div class="card shadow-sm">
            {% if product.image_url %}
            <img src="{{ product.image_url }}" class="card-img-top" style="height:250px;object-fit:cover;" alt="{{ product.name }}">
            {% endif %}
            <div class="card-body">
                <h4>{{ product.name }}</h4>
                <p>{{ product.description or '暂无描述' }}</p>
                <p class="text-danger fw-bold fs-4">{{ product.points_required }} 积分/件</p>
                <p class="text-muted">我的积分：<strong>{{ user.points }}</strong></p>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card shadow-sm">
            <div class="card-body">
                {% if error %}<div class="alert alert-danger">{{ error }}</div>{% endif %}
                <form method="POST">
                    <div class="mb-3">
                        <label class="form-label">兑换数量</label>
                        <input type="number" name="quantity" class="form-control" value="1" min="1"
                            {% if product.stock > 0 %}max="{{ product.stock }}"{% endif %} required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">收货人姓名</label>
                        <input type="text" name="contact_name" class="form-control" value="{{ session.user_name }}" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">联系电话</label>
                        <input type="text" name="contact_phone" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">收货地址</label>
                        <textarea name="address" class="form-control" rows="3" required></textarea>
                    </div>
                    <button type="submit" class="btn btn-danger w-100">确认兑换</button>
                </form>
            </div>
        </div>
    </div>
</div>
{% endif %}
{% endblock %}
```

---

### Task 11: 我的积分 + 我的订单

**Files:**
- Create: `templates/my_points.html`
- Create: `templates/my_orders.html`
- Modify: `app.py` — 添加路由

- [ ] **Step 1: 在 app.py 添加路由**

```python
@app.route('/my-points')
@login_required
def my_points():
    user = User.query.get(session['user_id'])
    logs = PointsLog.query.filter_by(user_id=user.id).order_by(PointsLog.created_at.desc()).limit(100).all()
    return render_template('my_points.html', user=user, logs=logs)


@app.route('/my-orders')
@login_required
def my_orders():
    orders = Order.query.filter_by(user_id=session['user_id']).order_by(Order.created_at.desc()).all()
    return render_template('my_orders.html', orders=orders)
```

- [ ] **Step 2: 创建 templates/my_points.html**

```html
{% extends "base.html" %}
{% block title %}我的积分{% endblock %}
{% block content %}
<h2 class="mb-4">我的积分</h2>
<div class="card shadow-sm mb-4">
    <div class="card-body text-center">
        <h6 class="text-muted">当前积分余额</h6>
        <h1 class="text-success display-3">{{ user.points }}</h1>
    </div>
</div>
<div class="card shadow-sm">
    <div class="card-header"><h5 class="mb-0">积分流水</h5></div>
    <div class="card-body p-0 table-responsive">
        <table class="table table-sm mb-0">
            <thead><tr><th>时间</th><th>来源</th><th>变动</th><th>余额</th></tr></thead>
            <tbody>
                {% for l in logs %}
                <tr>
                    <td>{{ l.created_at.strftime('%Y-%m-%d %H:%M') }}</td>
                    <td>{{ l.source }}</td>
                    <td class="{{ 'text-success' if l.points_change > 0 else 'text-danger' }}">
                        {{ '+' if l.points_change > 0 }}{{ l.points_change }}
                    </td>
                    <td>{{ l.balance_after }}</td>
                </tr>
                {% endfor %}
                {% if not logs %}
                <tr><td colspan="4" class="text-center text-muted py-3">暂无积分记录</td></tr>
                {% endif %}
            </tbody>
        </table>
    </div>
</div>
{% endblock %}
```

- [ ] **Step 3: 创建 templates/my_orders.html**

```html
{% extends "base.html" %}
{% block title %}我的订单{% endblock %}
{% block content %}
<h2 class="mb-4">我的订单</h2>
<div class="card shadow-sm">
    <div class="card-body p-0 table-responsive">
        <table class="table mb-0">
            <thead><tr><th>下单时间</th><th>商品</th><th>数量</th><th>消耗积分</th><th>状态</th></tr></thead>
            <tbody>
                {% for o in orders %}
                <tr>
                    <td>{{ o.created_at.strftime('%Y-%m-%d %H:%M') }}</td>
                    <td>{{ o.product.name if o.product else '已下架' }}</td>
                    <td>{{ o.quantity }}</td>
                    <td>{{ o.points_spent }}</td>
                    <td>
                        <span class="badge bg-{{ 'warning' if o.status=='待处理' else 'info' if o.status=='处理中' else 'primary' if o.status=='已发货' else 'success' }}">
                            {{ o.status }}
                        </span>
                    </td>
                </tr>
                {% endfor %}
                {% if not orders %}
                <tr><td colspan="5" class="text-center text-muted py-3">暂无订单</td></tr>
                {% endif %}
            </tbody>
        </table>
    </div>
</div>
{% endblock %}
```

---

### Task 12: 管理后台 - 商品管理

**Files:**
- Create: `templates/admin/products.html`
- Modify: `app.py` — 添加 `/admin/products` 路由

- [ ] **Step 1: 在 app.py 添加路由**

```python
@app.route('/admin/products', methods=['GET', 'POST'])
@admin_required
def admin_products():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            p = Product(
                name=request.form['name'],
                description=request.form.get('description', ''),
                image_url=request.form.get('image_url', ''),
                points_required=int(request.form.get('points_required', 0)),
                stock=int(request.form.get('stock', -1)),
                is_active=1
            )
            db.session.add(p)
            db.session.commit()
        elif action == 'edit':
            p = Product.query.get(int(request.form['id']))
            if p:
                p.name = request.form['name']
                p.description = request.form.get('description', '')
                p.image_url = request.form.get('image_url', '')
                p.points_required = int(request.form.get('points_required', 0))
                p.stock = int(request.form.get('stock', -1))
                db.session.commit()
        elif action == 'toggle':
            p = Product.query.get(int(request.form['id']))
            if p:
                p.is_active = 0 if p.is_active else 1
                db.session.commit()

    products = Product.query.order_by(Product.created_at.desc()).all()
    return render_template('admin/products.html', products=products)
```

- [ ] **Step 2: 创建 templates/admin/products.html**

```html
{% extends "base.html" %}
{% block title %}商品管理{% endblock %}
{% block content %}
<h2 class="mb-3">商品管理</h2>

<button class="btn btn-primary mb-3" data-bs-toggle="modal" data-bs-target="#productModal" onclick="clearForm()">添加商品</button>

<div class="card shadow-sm">
    <div class="card-body p-0 table-responsive">
        <table class="table mb-0">
            <thead><tr><th>图片</th><th>名称</th><th>所需积分</th><th>库存</th><th>状态</th><th>操作</th></tr></thead>
            <tbody>
                {% for p in products %}
                <tr>
                    <td>
                        {% if p.image_url %}<img src="{{ p.image_url }}" width="50" height="50" style="object-fit:cover;">{% else %}-{% endif %}
                    </td>
                    <td>{{ p.name }}</td>
                    <td>{{ p.points_required }}</td>
                    <td>{{ '无限' if p.stock == -1 else p.stock }}</td>
                    <td><span class="badge bg-{{ 'success' if p.is_active else 'secondary' }}">{{ '上架' if p.is_active else '下架' }}</span></td>
                    <td>
                        <button class="btn btn-sm btn-outline-primary" onclick="editProduct({{ p.id }}, '{{ p.name }}', '{{ p.description or '' }}', '{{ p.image_url or '' }}', {{ p.points_required }}, {{ p.stock }})">编辑</button>
                        <form method="POST" style="display:inline;">
                            <input type="hidden" name="action" value="toggle">
                            <input type="hidden" name="id" value="{{ p.id }}">
                            <button type="submit" class="btn btn-sm btn-outline-{{ 'warning' if p.is_active else 'success' }}">{{ '下架' if p.is_active else '上架' }}</button>
                        </form>
                    </td>
                </tr>
                {% endfor %}
                {% if not products %}
                <tr><td colspan="6" class="text-center text-muted py-3">暂无商品</td></tr>
                {% endif %}
            </tbody>
        </table>
    </div>
</div>

<!-- Modal for Add/Edit -->
<div class="modal fade" id="productModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <form method="POST">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalTitle">添加商品</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <input type="hidden" name="action" id="formAction" value="add">
                    <input type="hidden" name="id" id="productId" value="">
                    <div class="mb-3">
                        <label class="form-label">商品名称</label>
                        <input type="text" name="name" id="productName" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">描述</label>
                        <textarea name="description" id="productDesc" class="form-control" rows="2"></textarea>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">图片 URL</label>
                        <input type="url" name="image_url" id="productImage" class="form-control" placeholder="https://...">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">所需积分</label>
                        <input type="number" name="points_required" id="productPoints" class="form-control" required min="0">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">库存（-1 表示不限）</label>
                        <input type="number" name="stock" id="productStock" class="form-control" value="-1" min="-1">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
function clearForm() {
    document.getElementById('formAction').value = 'add';
    document.getElementById('productId').value = '';
    document.getElementById('productName').value = '';
    document.getElementById('productDesc').value = '';
    document.getElementById('productImage').value = '';
    document.getElementById('productPoints').value = '';
    document.getElementById('productStock').value = '-1';
    document.getElementById('modalTitle').textContent = '添加商品';
}
function editProduct(id, name, desc, image, points, stock) {
    document.getElementById('formAction').value = 'edit';
    document.getElementById('productId').value = id;
    document.getElementById('productName').value = name;
    document.getElementById('productDesc').value = desc;
    document.getElementById('productImage').value = image;
    document.getElementById('productPoints').value = points;
    document.getElementById('productStock').value = stock;
    document.getElementById('modalTitle').textContent = '编辑商品';
    new bootstrap.Modal(document.getElementById('productModal')).show();
}
</script>
{% endblock %}
```

---

### Task 13: 管理后台 - 订单管理

**Files:**
- Create: `templates/admin/orders.html`
- Modify: `app.py` — 添加 `/admin/orders` 路由

- [ ] **Step 1: 在 app.py 添加路由**

```python
@app.route('/admin/orders', methods=['GET', 'POST'])
@admin_required
def admin_orders():
    if request.method == 'POST':
        action = request.form.get('action')
        order_id = int(request.form.get('order_id'))
        order = Order.query.get(order_id)
        if order:
            if action == 'status':
                new_status = request.form.get('status')
                valid_statuses = ['待处理', '处理中', '已发货', '已完成']
                if new_status in valid_statuses:
                    order.status = new_status
                    db.session.commit()

    orders = Order.query.order_by(Order.created_at.desc()).all()
    return render_template('admin/orders.html', orders=orders)
```

- [ ] **Step 2: 创建 templates/admin/orders.html**

```html
{% extends "base.html" %}
{% block title %}订单管理{% endblock %}
{% block content %}
<h2 class="mb-4">订单管理</h2>
<div class="card shadow-sm">
    <div class="card-body p-0 table-responsive">
        <table class="table table-sm mb-0">
            <thead><tr><th>下单时间</th><th>用户</th><th>商品</th><th>数量</th><th>积分</th><th>收货人</th><th>联系电话</th><th>地址</th><th>状态</th><th>操作</th></tr></thead>
            <tbody>
                {% for o in orders %}
                <tr>
                    <td>{{ o.created_at.strftime('%m-%d %H:%M') }}</td>
                    <td>{{ o.user.name if o.user else '-' }}</td>
                    <td>{{ o.product.name if o.product else '已下架' }}</td>
                    <td>{{ o.quantity }}</td>
                    <td>{{ o.points_spent }}</td>
                    <td>{{ o.contact_name }}</td>
                    <td>{{ o.contact_phone }}</td>
                    <td>{{ o.address }}</td>
                    <td><span class="badge bg-{{ 'warning' if o.status=='待处理' else 'info' if o.status=='处理中' else 'primary' if o.status=='已发货' else 'success' }}">{{ o.status }}</span></td>
                    <td>
                        <form method="POST" class="d-flex gap-1">
                            <input type="hidden" name="action" value="status">
                            <input type="hidden" name="order_id" value="{{ o.id }}">
                            <select name="status" class="form-select form-select-sm" style="width:auto;">
                                {% for s in ['待处理', '处理中', '已发货', '已完成'] %}
                                <option value="{{ s }}" {{ 'selected' if o.status == s }}>{{ s }}</option>
                                {% endfor %}
                            </select>
                            <button type="submit" class="btn btn-sm btn-primary">更新</button>
                        </form>
                    </td>
                </tr>
                {% endfor %}
                {% if not orders %}
                <tr><td colspan="10" class="text-center text-muted py-3">暂无订单</td></tr>
                {% endif %}
            </tbody>
        </table>
    </div>
</div>
{% endblock %}
```

---

### Task 14: 部署配置

**Files:**
- Modify: `app.py` — 更新 SECRET_KEY 和数据库路径
- Create or verify: `Procfile`

- [ ] **Step 1: 确保 Procfile 正确**

```
web: gunicorn app:app
```

- [ ] **Step 2: 在 app.py 中添加 Render 兼容配置**

确保 app.py 末尾：

```python
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
```

- [ ] **Step 3: 部署到 Render**

1. 初始化 git 仓库：
```bash
git init
git add .
git commit -m "Initial commit"
```

2. 推送到 GitHub
3. 在 Render.com 创建新 Web Service，连接仓库
4. Build Command: `pip install -r requirements.txt`
5. Start Command: `gunicorn app:app`
6. 设置环境变量 `SECRET_KEY` 为随机字符串

---

### Task 15: 最终集成测试

- [ ] **Step 1: 启动应用并测试全部流程**

```bash
python app.py
```

验证以下流程：
1. 注册一个用户（非管理员名）
2. 注册何莎山 -> 确认显示 (管理员)
3. 管理员导入 `东4月.xlsx`
4. 查看首页仪表盘 -> 确认数据正确
5. 查看排名 -> 切换维度 -> 导出 Excel
6. 管理员添加商品
7. 普通用户兑换商品 -> 填写收货信息 -> 确认扣分
8. 管理员查看订单 -> 更新订单状态
9. 用户查看我的积分 -> 确认流水正确
10. 用户查看我的订单 -> 确认状态更新
```

