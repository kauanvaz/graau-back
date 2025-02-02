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
    # sharepoint_data = PaneisMetas().get_monitoramento(item_id=3868)

    sharepoint_data = [
        {
            "unidades_fiscalizadas": "P. M. CIDADE",
            "n_processo_eTCE": "TC/XXXXXX/20XX",
            "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
            "exercicios": "20XX, 20YY",
            "VRF": "R$ 100.000,00"
        }
    ]
    
    # generate_json(sharepoint_data)
    
    # Criar o gerador de relatórios
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    
    # Gerar relatório
    success = generator.generate_report(
        context=sharepoint_data[0],
        output_path="src/relatorios/report-template.docx",
        cover_image_path="src/cover_images/cover_page_1.jpg"
    )
    
    if success:
        print("Relatório gerado com sucesso.")
    else:
        print("Erro ao gerar relatório.")