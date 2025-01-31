import json
from src.sharepoint import PaneisMetas
from src.report_generator import ReportGenerator
from pathlib import Path

def generate_json(data):
    # Salvar os dados em um arquivo JSON
    
    filename = "test_data.json"
    
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    print(f"Arquivo {filename} gerado.")

if __name__ == "__main__":
    sharepoint_data = PaneisMetas().get_monitoramento(item_id=3868)
    
    generate_json(sharepoint_data)
    
    # Criar o gerador de relatórios
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    
    # Gerar relatório
    result_report_path = Path('src/relatorios')
    success = generator.generate_report(
        context=sharepoint_data[0],
        output_path="src/relatorios/report-template.docx"
    )
    
    print("Arquivo src/relatorios/report-template.docx gerado.")