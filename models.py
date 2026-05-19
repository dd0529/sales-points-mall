from extensions import db
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
    opening_date = db.Column(db.String(50))
    biz_org_name = db.Column(db.String(200))
    product_name = db.Column(db.String(100))
    salesperson = db.Column(db.String(50))
    salesperson_code = db.Column(db.String(50))
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
