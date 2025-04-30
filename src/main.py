import json
from sharepoint import Sharepoint
from report_generator import ReportGenerator
from utils import get_status_processo, format_data , load_json

def generate_json(data):
    # Salvar os dados em um arquivo JSON
    
    filename = "test_data.json"
    
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    print(f"Arquivo {filename} gerado.")

if __name__ == "__main__":
    sharepoint_data = Sharepoint().get_acao_controle_data(item_id=4102)
    
    generate_json(sharepoint_data)
    
    data = sharepoint_data[0]
    data["seccoes"] = load_json("examples/sections.json")
    
    tipo_relatorio = "Recurso" # Vem do frontend
    data["tipo_relatorio"] = tipo_relatorio
    
    formatted_data = format_data(data)
    
    status_processo = get_status_processo(tipo_relatorio, formatted_data["processo_tipo"])
    
    if status_processo:
        formatted_data["status_processo"] = status_processo
    
    # Criar o gerador de relatórios
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    
    # Gerar relatório
    success = generator.generate_report(
        context=formatted_data,
        output_path="src/reports/result.docx",
        cover_image_path="src/cover_images/cover_page_1.jpg"
    )
    
    if success:
        print("Relatório gerado com sucesso.")
    else:
        print("Erro ao gerar relatório.")
    