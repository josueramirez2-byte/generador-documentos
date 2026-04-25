import os
import io
import zipfile
import json
import re
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

        doc = DocxTemplate(path)
        campos = doc.get_undeclared_template_variables()
        return render_template('rellenar.html', campos=campos, nombre_archivo=nombre_archivo)
    except Exception as e:
        return f"Error al abrir plantilla: {str(e)}", 500

@app.route('/generar_final', methods=['POST'])
def generar_final():
    try:
        nombre_archivo_base = request.form.get('nombre_archivo')
        patron_nombre = request.form.get('nombre_dinamico')
        path_plantilla = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo_base)
        datos_raw = request.form.get('datos_totales')
        
        datos_lote = json.loads(datos_raw)

        def limpiar_nombre(n):
            # Elimina caracteres no permitidos en nombres de archivos
            return re.sub(r'[\\/*?:"<>|]', "", n)

        def procesar_nombre(patron, datos, i):
            resultado = patron
            for k, v in datos.items():
            # Esto busca la llave sin importar si es Mayúscula o Minúscula
                pattern = re.compile(re.escape(f"{{{{{k}}}}}"), re.IGNORECASE)
                resultado = pattern.sub(str(v), resultado)
            if resultado == patron and len(datos_lote) > 1:
                resultado = f"{resultado}_{i+1}"
            return limpiar_nombre(resultado)

        if len(datos_lote) == 1:
            doc = DocxTemplate(path_plantilla)
            doc.render(datos_lote[0])
            nombre_descarga = procesar_nombre(patron_nombre, datos_lote[0], 0) + ".docx"
            
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            return send_file(output, as_attachment=True, download_name=nombre_descarga)
        
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                for i, datos in enumerate(datos_lote):
                    doc = DocxTemplate(path_plantilla)
                    doc.render(datos)
                    nombre_doc = procesar_nombre(patron_nombre, datos, i) + ".docx"
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    zf.writestr(nombre_doc, doc_io.getvalue())
            zip_buffer.seek(0)
            return send_file(zip_buffer, as_attachment=True, download_name="lote_documentos.zip")

    except Exception as e:
        return f"Error en generación: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
