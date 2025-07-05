from flask import Flask, request, render_template, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
import os
from .xlsb_parser import XLSBParser

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB

db = SQLAlchemy(app)

class XLSBFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    path = db.Column(db.String, unique=True, nullable=False)
    original_name = db.Column(db.String, nullable=False)

class DataRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey('XLSB_file.id'), nullable=False)
    table = db.Column(db.String, nullable=False)
    id2 = db.Column(db.String)
    key = db.Column(db.String)
    value = db.Column(db.String)

    file = db.relationship('XLSBFile', backref=db.backref('records', lazy=True))

def init_db():
    """Create database and upload folder if they do not exist."""
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    db.create_all()

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        uploaded = request.files.getlist('files')
        for f in uploaded:
            filename = f.filename
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            f.save(save_path)
            xlsb = XLSBFile(path=save_path, original_name=filename)
            db.session.add(xlsb)
            db.session.commit()
            parser = XLSBParser(save_path)
            for table_name, df in parser.parse_all().items():
                for _, row in df.iterrows():
                    rec = DataRecord(file_id=xlsb.id, table=table_name,
                                     id2=row.get('Id2'), key=row['Key'], value=str(row['Value']))
                    db.session.add(rec)
        db.session.commit()
        return redirect(url_for('dashboard'))

    return render_template('upload.html')

@app.route('/dashboard')
def dashboard():
    """Display parsed data with optional filtering by table or file."""
    table_filter = request.args.get('table')
    file_filter = request.args.get('file')
    id2_filter = request.args.get('id2')

    query = DataRecord.query
    if table_filter:
        query = query.filter_by(table=table_filter)
    if file_filter:
        query = query.filter_by(file_id=file_filter)
    if id2_filter:
        query = query.filter_by(id2=id2_filter)
    records = query.order_by(DataRecord.table).all()

    grouped = {}
    for r in records:
        grouped.setdefault(r.table, []).append(r)

    parts = []
    for table_name, rows in grouped.items():
        row_html = ''.join(
            '<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>'.format(
                r.file.original_name, r.id2 or '', r.key, r.value) for r in rows
        )
        parts.append(
            '<h2>{}</h2>'.format(table_name) +
            '<table border="1">'
            '<tr><th>File</th><th>Id2</th><th>Key</th><th>Value</th></tr>' +
            row_html +
            '</table>'
        )

    html_tables = '\n'.join(parts)

    return render_template('dashboard.html', tables=html_tables)

@app.route('/files')
def list_files():
    files = XLSBFile.query.all()
    return render_template('files.html', files=files)

@app.route('/delete/<int:file_id>')
def delete_file(file_id):
    xlsb = XLSBFile.query.get_or_404(file_id)
    DataRecord.query.filter_by(file_id=xlsb.id).delete()
    if os.path.isfile(xlsb.path):
        os.remove(xlsb.path)
    db.session.delete(xlsb)
    db.session.commit()
    return redirect(url_for('list_files'))

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
