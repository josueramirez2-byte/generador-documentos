import os
from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Crear la carpeta de subidas si no existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('subir.html')

@app.route('/analizar_plantilla', methods=['POST'])
def analizar_plantilla():
    if 'archivo' not in request.files:
        return "No se seleccionó ningún archivo", 400
    
    file = request.files['archivo']
    if file.filename == '':
        return "Nombre de archivo vacío", 400

    path_plantilla = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(path_plantilla)

    # Extraer variables del Word {{ variable }}
    doc = DocxTemplate(path_plantilla)
    campos_detectados = doc.get_undeclared_template_variables()
    
    return render_template('rellenar.html', campos=campos_detectados, nombre_archivo=file.filename)

@app.route('/generar_documento', methods=['POST'])
def generar_documento():
    nombre_archivo = request.form.get('nombre_archivo')
    path_plantilla = os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo)
    
    doc = DocxTemplate(path_plantilla)
    
    # Recoger datos del formulario
    datos = {key: value for key, value in request.form.items() if key != 'nombre_archivo'}
    
    doc.render(datos)
    
    # Guardar en memoria para descarga
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    
    return send_file(
        target_stream,
        as_attachment=True,
        download_name=f"FINAL_{nombre_archivo}",
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    # Configuración para Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)