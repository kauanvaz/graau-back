import json
from sharepoint import Sharepoint
from report_generator import ReportGenerator
from utils import format_data, get_status_processo

def generate_json(data):
    # Salvar os dados em um arquivo JSON
    
    filename = "test_data.json"
    
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    print(f"Arquivo {filename} gerado.")

if __name__ == "__main__":
    sharepoint_data = Sharepoint().get_acao_controle_data(item_id=4102)
    
    generate_json(sharepoint_data)
    
    # Criar o gerador de relatórios
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    formatted_data = format_data(sharepoint_data[0])
    formatted_data["seccoes"] = [
        {
            "title": "Introdução",
            "subtitles": [
                {
                    "title": "Objetivo",
                    "subtitles": [
                        {
                            "title": "Objetivo Amplo",
                            "subtitles": []
                        },
                        {
                            "title": "Objetivo Específico",
                            "subtitles": []
                        }
                    ]
                },
                {
                    "title": "Escopo",
                    "subtitles": [
                        {
                            "title": "Escopo geral",
                            "subtitles": []
                        },
                        {
                            "title": "Escopo intrínseco",
                            "subtitles": [
                                {
                                    "title": "Jurisdição", "subtitles": []
                                }
                            ]
                        }
                    ]
                }
            ]
        },
        {
            "title": "Metodologia",
            "subtitles": [
                {
                    "title": "Coleta de Dados",
                    "subtitles": []
                },
                {
                    "title": "Análise estatística",
                    "subtitles": []
                }
            ]
        },
        {
            "title": "Resultados",
            "subtitles": [
                {
                    "title": "Resultados gerais",
                    "subtitles": []
                },
                {
                    "title": "Resultados específicos",
                    "subtitles": []
                }
            ]
        },
        {
            "title": "Conclusão",
            "subtitles": [
                {
                    "title": "Considerações finais",
                    "subtitles": []
                },
                {
                    "title": "Recomendações",
                    "subtitles": []
                }
            ]
        }
    ]
    
    tipo_relatorio = "Recurso" # Vem do frontend
    formatted_data["tipo_relatorio"] = tipo_relatorio
    status_processo = get_status_processo(tipo_relatorio, formatted_data["processo_tipo"])
    
    if status_processo:
        formatted_data["status_processo"] = status_processo
    
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
    