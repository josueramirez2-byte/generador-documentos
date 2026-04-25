import os
import io
import zipfile
import json
from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate

app = Flask(__name__)

# Configuración de ruta absoluta para Render
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(basedir, 'uploads')

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    try:
        plantillas = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith('.docx')]
    except Exception:
        plantillas = []
    return render_template('subir.html', plantillas=plantillas)

@app.route('/analizar_plantilla', methods=['POST'])
def analizar_plantilla():
    try:
        plantilla_sel = request.form.get('plantilla_existente')
        file = request.files.get('archivo')

        if file and file.filename != '':
            nombre_archivo = file.filename
            path = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
            file.save(path)
        elif plantilla_sel:
            nombre_archivo = plantilla_sel
            path = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
        else:
            return "Error: No se seleccionó ninguna plantilla", 400

        # Intentar abrir el Word
        doc = DocxTemplate(path)
        campos = doc.get_undeclared_template_variables()
        return render_template('rellenar.html', campos=campos, nombre_archivo=nombre_archivo)
    
    except Exception as e:
        # Esto nos dirá el error real en la pantalla
        return f"Error al leer el Word: {str(e)}. Asegúrate de que el archivo no tenga errores de sintaxis como {{{{ campo } (falta una llave).", 500

@app.route('/generar_final', methods=['POST'])
def generar_final():
    try:
        nombre_archivo = request.form.get('nombre_archivo')
        path_plantilla = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
        datos_raw = request.form.get('datos_totales')
        
        if not datos_raw:
            return "Error: No se recibieron datos", 400
            
        datos_lote = json.loads(datos_raw)

        if len(datos_lote) == 1:
            doc = DocxTemplate(path_plantilla)
            doc.render(datos_lote[0])
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            return send_file(output, as_attachment=True, 
                             download_name=f"generado_{nombre_archivo}",
                             mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
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
    except Exception as e:
        return f"Error al generar: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
