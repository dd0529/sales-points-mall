import os
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from extensions import db
from sqlalchemy import func, desc
from werkzeug.utils import secure_filename
from utils import parse_excel

app = Flask(__name__)
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(basedir, "data.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

db.init_app(app)

# models must be imported after db init
from models import User, SalesRecord, Product, Order, PointsLog, ImportLog
from auth import login_required, admin_required, ADMIN_NAMES

@app.context_processor
def utility_processor():
    return dict(enumerate=enumerate)

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


@app.route('/ranking')
def ranking():
    dim = request.args.get('dim', 'salesperson')
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


@app.route('/mall')
@login_required
def mall():
    products = Product.query.filter_by(is_active=1).order_by(Product.points_required).all()
    user = User.query.get(session['user_id'])
    return render_template('mall.html', products=products, user=user)


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


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
