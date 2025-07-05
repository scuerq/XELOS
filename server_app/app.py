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
    __tablename__ = 'XLSB_file'
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
            if not filename:
                continue  # Ignore les fichiers sans nom pour éviter IsADirectoryError
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            f.save(save_path)
            xlsb = XLSBFile(path=save_path, original_name=filename)
            db.session.add(xlsb)
            db.session.commit()
            print("UPLOAD: appel parser.parse_all() sur", save_path)
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
    """Display parsed data with optional filtering by table or file, groupées par nom de table."""
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

    # Sections attendues
    section_names = ["SYNTHESE", "ANALYSE_PRP", "ANALYSE_FINANCEMENT", "ANALYSE_LOYERS"]
    tables = {name: [] for name in section_names}
    for r in records:
        if r.table in tables:
            tables[r.table].append(r)
        else:
            # Ajoute les tables non prévues à la fin
            tables.setdefault(r.table, []).append(r)
    return render_template('dashboard.html', tables=tables, section_names=section_names)

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
