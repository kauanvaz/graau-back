from typing import Dict, Union, Any
from docxtpl import DocxTemplate
from pathlib import Path
import logging

class ReportGenerator:
    def __init__(self, template_path: Union[str, Path]):
        """
        Inicializa o gerador de relatórios com um template
        
        Args:
            template_path: Caminho para o arquivo de template .docx
        """
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template não encontrado: {template_path}")
            
        self.doc = DocxTemplate(str(self.template_path))
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

    def generate_report(self, 
                       context: Dict[str, Any],
                       output_path: Union[str, Path]) -> bool:
        """
        Gera o relatório final substituindo as variáveis do template
        
        Args:
            context: Dicionário com os dados para preencher o template
            output_path: Caminho para salvar o arquivo
            
        Returns:
            bool: True se o relatório foi gerado com sucesso
        """
        try:
            # Renderizar o template com o contexto fornecido
            self.doc.render(context)
            
            # Salvar o documento
            output_path = Path(output_path)
            self.doc.save(str(output_path))
            
            self.logger.info(f"Relatório gerado com sucesso: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao gerar relatório: {str(e)}")
            return False

# Exemplo de uso
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    # Dados do SharePoint (exemplo)
    sharepoint_data = [
        {
            "unidades_fiscalizadas": "P. M. CIDADE",
            "n_processo_eTCE": "TC/XXXXXX/20XX",
            "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
            "exercicios": "20XX, 20YY",
            "VRF": "R$ 100.000,00"
        }
    ]
    
    # Criar o gerador de relatórios
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    
    # Gerar relatório
    success = generator.generate_report(
        context=sharepoint_data[0],
        output_path="src/relatorios/report_example.docx"
    )
