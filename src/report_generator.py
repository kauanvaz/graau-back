from typing import Dict, Union, Any, Optional
from docx.shared import Pt, Inches
from docxtpl import DocxTemplate
from pathlib import Path
import logging

class ReportGenerator:
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

    def add_cover_page(self, doc: DocxTemplate, image_path: Union[str, Path]) -> bool:
        """
        Adds a full-page cover image to the document before the template content.
        
        Args:
            doc: DocxTemplate instance
            image_path: Path to the cover image
            
        Returns:
            bool: True if the image was added successfully
        """
        try:
            image_path = Path(image_path)
            if not image_path.exists():
                raise FileNotFoundError(f"Image not found: {image_path}")

            # Get the underlying docx document
            document = doc.get_docx()
            
            # Get the first section
            section = document.sections[0]
            
            # Remove margins for the cover page
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)
            
            # Insert a new paragraph at the beginning of the document
            paragraph = document.paragraphs[0]
            run = paragraph.add_run()
            
            # Configure paragraph format
            paragraph_format = paragraph.paragraph_format
            paragraph_format.left_indent = Inches(0)
            paragraph_format.right_indent = Inches(0)
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)
            paragraph_format.line_spacing = 1.0

            # Add the image with page dimensions
            width = section.page_width
            height = section.page_height
            run.add_picture(str(image_path), width=width, height=height)

            # Add a new section with normal margins for the template content
            document.add_section()
            new_section = document.sections[-1]
            new_section.left_margin = Inches(1)
            new_section.right_margin = Inches(1)
            new_section.top_margin = Inches(1)
            new_section.bottom_margin = Inches(1)

            return True
            
        except Exception as e:
            self.logger.error(f"Error adding cover image: {str(e)}")
            return False

    def generate_report(self, context: dict, output_path: str, cover_image_path: Optional[Union[str, Path]] = None, graau_params: dict = {}) -> bool:
        """
        Generates the report using the template and provided context.
        
        Args:
            context: Dictionary with the template context
            output_path: Path to save the output file
            cover_image_path: Optional path to the cover image
            graau_params: Dictionary with more information from the user to the report
        Returns:
            bool: True if the report was generated successfully
        """
        try:
            # Create a new instance of DocxTemplate for each report generation
            doc = DocxTemplate(self.template_path)
            
            if cover_image_path:
                if not self.add_cover_page(doc, cover_image_path):
                    return False
            
            # Render the template with the context
            doc.render(context)
            
            # Save the document
            doc.save(output_path)
            
            self.logger.info(f"Report generated successfully: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error generating report: {str(e)}")
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
        output_path="src/reports/report_example.docx",
        cover_image_path="src/cover_images/cover_page_1.jpg"
    )
