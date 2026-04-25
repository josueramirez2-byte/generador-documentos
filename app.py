import os
import io
import zipfile
import json
from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    # Listar plantillas para el selector
    plantillas = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith('.docx')]
    return render_template('subir.html', plantillas=plantillas)

@app.route('/analizar_plantilla', methods=['POST'])
def analizar_plantilla():
    plantilla_sel = request.form.get('plantilla_existente')
    file = request.files.get('archivo')

    if file and file.filename != '':
        nombre_archivo = file.filename
        path = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
        file.save(path)
    elif plantilla_sel:
        nombre_archivo = plantilla_sel
    else:
        return "Error: Selecciona o sube una plantilla", 400

    doc = DocxTemplate(os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo))
    campos = doc.get_undeclared_template_variables()
    return render_template('rellenar.html', campos=campos, nombre_archivo=nombre_archivo)

@app.route('/generar_final', methods=['POST'])
def generar_final():
    nombre_archivo = request.form.get('nombre_archivo')
    path_plantilla = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
    datos_lote = json.loads(request.form.get('datos_totales'))

    if not datos_lote:
        return "No hay datos", 400

    # CASO 1: Solo un registro -> Descargar .docx directamente
    if len(datos_lote) == 1:
        doc = DocxTemplate(path_plantilla)
        doc.render(datos_lote[0])
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, 
                         download_name=f"generado_{nombre_archivo}",
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    # CASO 2: Varios registros -> Descargar .zip
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            for i, datos in enumerate(datos_lote):
                doc = DocxTemplate(path_plantilla)
                doc.render(datos)
                doc_io = io.BytesIO()
                doc.save(doc_io)
                zf.writestr(f"documento_{i+1}.docx", doc_io.getvalue())
        
        zip_buffer.seek(0)
        return send_file(zip_buffer, as_attachment=True, 
                         download_name="lote_documentos.zip", 
                         mimetype='application/zip')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
