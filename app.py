# app.py
from flask import Flask, request, jsonify, send_file
import os
import json
import uuid
import datetime
import threading
import time
from src.report_generator import ReportGenerator
from src.sharepoint import Sharepoint
import logging
import concurrent.futures

# Configuração do logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(lineno)d - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("api.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Diretório onde os relatórios serão armazenados
REPORTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src/reports')
if not os.path.exists(REPORTS_DIR):
    os.makedirs(REPORTS_DIR)

# Diretório para armazenar relatórios em andamento
PENDING_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pending_reports')
if not os.path.exists(PENDING_DIR):
    os.makedirs(PENDING_DIR)

# Diretório para armazenar imagens de capa temporárias
COVER_IMAGES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src/temp_cover_images')
if not os.path.exists(COVER_IMAGES_DIR):
    os.makedirs(COVER_IMAGES_DIR)

# Arquivo para rastrear os relatórios
REPORTS_TRACKER = os.path.join(REPORTS_DIR, 'reports_tracker.json')
if not os.path.exists(REPORTS_TRACKER):
    with open(REPORTS_TRACKER, 'w') as f:
        json.dump([], f)

# Configurações da API - TORNADAS FACILMENTE CONFIGURÁVEIS
app.config['REPORT_EXPIRATION_MINUTES'] = 1  # Valor padrão de 30 minutos
app.config['CLEANUP_INTERVAL_SECONDS'] = 60  # Verificar a cada 5 minutos
app.config['ALLOWED_IMAGE_EXTENSIONS'] = {'png', 'jpg', 'jpeg'}
app.config['MAX_IMAGE_SIZE'] = 5 * 1024 * 1024  # 5MB

# Pool de threads para processamento assíncrono
executor = concurrent.futures.ThreadPoolExecutor(max_workers=5)
# Dicionário para rastrear tarefas assíncronas
tasks = {}


def load_reports_tracker():
    """Carrega o rastreador de relatórios do arquivo JSON."""
    try:
        with open(REPORTS_TRACKER, 'r') as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return []


def save_reports_tracker(tracker_data):
    """Salva o rastreador de relatórios no arquivo JSON."""
    with open(REPORTS_TRACKER, 'w') as f:
        json.dump(tracker_data, f, indent=4)

def allowed_file(filename):
    """Verifica se o arquivo possui uma extensão permitida."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_IMAGE_EXTENSIONS']

def cleanup_old_reports():
    """Remove relatórios antigos com base na configuração REPORT_EXPIRATION_MINUTES."""
    while True:
        logger.info(f"Iniciando limpeza de relatórios antigos (limite: {app.config['REPORT_EXPIRATION_MINUTES']} minutos)")
        tracker_data = load_reports_tracker()
        current_time = datetime.datetime.now()
        updated_tracker = []
        
        for report in tracker_data:
            created_date = datetime.datetime.fromisoformat(report['created_at'])
            age_minutes = (current_time - created_date).total_seconds() / 60
            
            if age_minutes > app.config['REPORT_EXPIRATION_MINUTES']:
                report_path = os.path.join(REPORTS_DIR, report['filename'])
                try:
                    if os.path.exists(report_path):
                        os.remove(report_path)
                        logger.info(f"Relatório excluído: {report['filename']}")
                        
                        # Remover imagem de capa associada, se houver
                    if 'cover_image' in report and report['cover_image']:
                        cover_path = os.path.join(COVER_IMAGES_DIR, report['cover_image'])
                        if os.path.exists(cover_path):
                            os.remove(cover_path)
                            logger.info(f"Imagem de capa excluída: {report['cover_image']}")
                            
                    # Remover arquivo de status
                    status_file = os.path.join(PENDING_DIR, f"{report.get('task_id', '')}.json")
                    if os.path.exists(status_file):
                        os.remove(status_file)
                        
                except Exception as e:
                    logger.error(f"Erro ao excluir relatório {report['filename']}: {str(e)}")
            else:
                updated_tracker.append(report)
        
        save_reports_tracker(updated_tracker)
        logger.info(f"Limpeza concluída. {len(tracker_data) - len(updated_tracker)} relatórios removidos.")
        
        # Executar a limpeza com base no intervalo configurado
        time.sleep(app.config['CLEANUP_INTERVAL_SECONDS'])


# Iniciar o thread de limpeza ao iniciar a aplicação
cleanup_thread = threading.Thread(target=cleanup_old_reports, daemon=True)
cleanup_thread.start()


def generate_report_task(data, filepath, task_id, cover_image_path):
    """Função para gerar o relatório de forma assíncrona."""
    try:
        logger.info(f"Iniciando geração assíncrona do relatório: {task_id}")
        
        # Criar um arquivo de status para acompanhamento
        status_file = os.path.join(PENDING_DIR, f"{task_id}.json")
        with open(status_file, 'w') as f:
            json.dump({
                "status": "processing",
                "message": "Obtendo dados do SharePoint",
                "progress": 10,
                "created_at": datetime.datetime.now().isoformat()
            }, f, indent=4)
        
        # Obter dados do SharePoint
        sharepoint = Sharepoint()
        sharepoint_data = sharepoint.get_acao_controle_data(item_id=data['sharepoint_id'])
        
        # Atualizar status
        with open(status_file, 'w') as f:
            json.dump({
                "status": "processing",
                "message": "Gerando relatório com os dados obtidos",
                "progress": 50,
                "created_at": datetime.datetime.now().isoformat()
            }, f, indent=4)
        
        # Gerar relatório
        report_generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
        
        # Adicionar parâmetros para o relatório
        report_params = data['report_params'].copy() if 'report_params' in data else {}
        
        report_generator.generate_report(
            context=sharepoint_data[0],
            output_path=filepath,
            cover_image_path=cover_image_path,
            graau_params=report_params
        )
        
        # Registrar o relatório no rastreador
        report_info = {
            "filename": os.path.basename(filepath),
            "created_at": datetime.datetime.now().isoformat(),
            "parameters": {
                "sharepoint_id": data['sharepoint_id'],
                "report_params": data['report_params']
            },
            "status": "completed",
            "task_id": task_id,
            "cover_image": os.path.basename(cover_image_path) if cover_image_path else None
        }
        
        tracker_data = load_reports_tracker()
        tracker_data.append(report_info)
        save_reports_tracker(tracker_data)
        
        # Atualizar status final
        with open(status_file, 'w') as f:
            json.dump({
                "status": "completed",
                "message": "Relatório gerado com sucesso",
                "progress": 100,
                "created_at": datetime.datetime.now().isoformat(),
                "filename": os.path.basename(filepath)
            }, f, indent=4)
        
        logger.info(f"Relatório gerado com sucesso: {task_id}")
        return report_info
        
    except Exception as e:
        error_message = f"Erro ao gerar relatório: {str(e)}"
        logger.error(error_message)
        
        # Atualizar status com erro
        with open(status_file, 'w') as f:
            json.dump({
                "status": "error",
                "message": error_message,
                "progress": 0,
                "created_at": datetime.datetime.now().isoformat()
            }, f, indent=4)
        
        return {"error": error_message}

@app.route('/api/upload-cover-image', methods=['POST'])
def upload_cover_image():
    """Endpoint para fazer upload de uma imagem de capa."""
    try:
        # Verificar se a requisição contém um arquivo
        if 'file' not in request.files:
            return jsonify({"error": "Nenhum arquivo enviado"}), 400
            
        file = request.files['file']
        
        # Verificar se um arquivo foi selecionado
        if file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400
            
        # Verificar extensão do arquivo
        if not allowed_file(file.filename):
            return jsonify({
                "error": f"Formato de arquivo não permitido. Use: {', '.join(app.config['ALLOWED_IMAGE_EXTENSIONS'])}"
            }), 400
            
        # Verificar tamanho do arquivo
        if len(file.read()) > app.config['MAX_IMAGE_SIZE']:
            file.seek(0)  # Reset do ponteiro do arquivo
            return jsonify({"error": f"Arquivo muito grande. Tamanho máximo: {app.config['MAX_IMAGE_SIZE'] / 1024 / 1024}MB"}), 400
            
        file.seek(0)  # Reset do ponteiro do arquivo
        
        # Gerar nome único para o arquivo
        image_id = str(uuid.uuid4())
        extension = file.filename.rsplit('.', 1)[1].lower()
        filename = f"cover_{image_id}.{extension}"
        filepath = os.path.join(COVER_IMAGES_DIR, filename)
        
        # Salvar o arquivo
        file.save(filepath)
        
        return jsonify({
            "success": True,
            "message": "Imagem de capa enviada com sucesso",
            "image_id": image_id,
            "filename": filename
        }), 201
        
    except Exception as e:
        logger.error(f"Erro ao fazer upload da imagem: {str(e)}")
        return jsonify({"error": f"Erro ao processar upload: {str(e)}"}), 500

@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """
    Endpoint para gerar um relatório baseado em dados JSON (de forma assíncrona).
    Espera receber um JSON com:
    - sharepoint_id: Parâmetros para buscar dados no SharePoint
    - report_params: Parâmetros para gerar o relatório
    - cover_image_id: ID da imagem de capa previamente enviada
    """
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Nenhum dado JSON recebido"}), 400
        
        # Validação de campos obrigatórios
        required_fields = ['sharepoint_id', 'report_params', 'cover_image_id']
        missing_fields = [field for field in required_fields if field not in data]
        if missing_fields:
            return jsonify({"error": f"Campos obrigatórios ausentes: {', '.join(missing_fields)}"}), 400
        
        # Verificar a imagem de capa, se informada
        cover_image_path = None
        if 'cover_image_id' in data and data['cover_image_id']:
            # Procurar a imagem pelo ID
            for filename in os.listdir(COVER_IMAGES_DIR):
                if data['cover_image_id'] in filename:
                    cover_image_path = os.path.join(COVER_IMAGES_DIR, filename)
                    break
            
        if not cover_image_path or not os.path.exists(cover_image_path):
            return jsonify({"error": "Imagem de capa não encontrada. Faça o upload novamente."}), 404
        
        # Gerar nome de arquivo único
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        task_id = str(uuid.uuid4())
        filename = f"{'relatorio'}_{timestamp}_{task_id[:8]}.docx"
        filepath = os.path.join(REPORTS_DIR, filename)
        
        # Iniciar geração de relatório em thread separada
        future = executor.submit(generate_report_task, data, filepath, task_id, cover_image_path)
        tasks[task_id] = future
        
        # Retornar imediatamente com o ID da tarefa
        return jsonify({
            "success": True,
            "message": "Geração de relatório iniciada",
            "task_id": task_id,
            "status": "processing"
        }), 202
        
    except Exception as e:
        logger.error(f"Erro ao iniciar geração de relatório: {str(e)}")
        return jsonify({"error": f"Erro ao iniciar geração de relatório: {str(e)}"}), 500


@app.route('/api/report-status/<task_id>', methods=['GET'])
def get_report_status(task_id):
    """Verifica o status de uma tarefa de geração de relatório."""
    try:
        status_file = os.path.join(PENDING_DIR, f"{task_id}.json")
        
        if os.path.exists(status_file):
            with open(status_file, 'r') as f:
                status_data = json.load(f)
                
            # Se estiver completo, adicionar link para download
            if status_data["status"] == "completed":
                tracker_data = load_reports_tracker()
                for report in tracker_data:
                    if report.get("task_id") == task_id:
                        status_data["download_url"] = f"/api/reports/{task_id}"
                        status_data["filename"] = report["filename"]
                        break
            
            return jsonify(status_data)
        
        # Verificar se está no tracker (processamento pode ter terminado, mas ainda não removido)
        tracker_data = load_reports_tracker()
        for report in tracker_data:
            if report.get("task_id") == task_id:
                return jsonify({
                    "status": "completed",
                    "message": "Relatório gerado com sucesso",
                    "progress": 100,
                    "download_url": f"/api/reports/{task_id}",
                    "filename": report["filename"]
                })
        
        # Não encontrado
        return jsonify({"error": "Tarefa não encontrada"}), 404
        
    except Exception as e:
        logger.error(f"Erro ao verificar status da tarefa {task_id}: {str(e)}")
        return jsonify({"error": f"Erro ao verificar status: {str(e)}"}), 500

@app.route('/api/reports/<report_id>', methods=['GET'])
def download_report(report_id):
    """Download de um relatório específico com base no report_id."""
    tracker_data = load_reports_tracker()
    
    # Encontrar o relatório pelo ID da tarefa
    report = None
    for r in tracker_data:
        if report_id == r.get("task_id") or report_id in r['filename']:
            report = r
            break
    
    if not report:
        return jsonify({"error": "Relatório não encontrado"}), 404
    
    filepath = os.path.join(REPORTS_DIR, report['filename'])
    if not os.path.exists(filepath):
        return jsonify({"error": "Arquivo de relatório não encontrado no servidor"}), 404
    
    return send_file(
        filepath, 
        as_attachment=True,
        download_name=f"{report['filename']}.docx",
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)