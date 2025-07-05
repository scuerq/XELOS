from flask import Flask, request, render_template_string, redirect, url_for
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

    return render_template_string('''
        <h1>Upload XLSB Files</h1>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple>
            <input type="submit" value="Upload">
        </form>
    ''')

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
    records = query.all()

    rows = [
        '<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>'.format(
            r.file.original_name, r.table, r.id2 or '', r.key, r.value)
        for r in records
    ]

    html_table = (
        '<table border="1">'
        '<tr><th>File</th><th>Table</th><th>Id2</th><th>Key</th><th>Value</th></tr>{}'
        '</table>'
    ).format(''.join(rows))

    form = (
        '<form method="get">'
        'Filter table: <input name="table" value="{0}"> '
        'File id: <input name="file" value="{1}"> '
        'Id2: <input name="id2" value="{2}"> '
        '<input type="submit" value="Filter">'
        '</form>'
    ).format(table_filter or '', file_filter or '', id2_filter or '')

    return render_template_string('<h1>Dashboard</h1>' + form + html_table)

@app.route('/files')
def list_files():
    files = XLSBFile.query.all()
    rows = [
        f'<tr><td>{f.id}</td><td>{f.original_name}</td></tr>'
        for f in files
    ]
    table = '<table border="1"><tr><th>ID</th><th>Name</th></tr>{}</table>'.format(''.join(rows))
    return render_template_string('<h1>Uploaded Files</h1>' + table)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
