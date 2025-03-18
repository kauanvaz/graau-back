from typing import Union, Optional
from docx.shared import Pt, Inches
from docxtpl import DocxTemplate
from pathlib import Path
import logging
import zipfile
import os
import shutil
import tempfile

class ReportGenerator:
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

    def add_cover_page(self, doc: DocxTemplate, image_path: Union[str, Path]) -> bool:
        """
        Adds a full-page cover image to the document before the template content.
        """
        try:
            image_path = Path(image_path)
            if not image_path.exists():
                raise FileNotFoundError(f"Image not found: {image_path}")

            document = doc.get_docx()
            section = document.sections[0]

            # Remove margins for the cover page
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)

            # Clear the first paragraph of any existing content
            if document.paragraphs:
                first_paragraph = document.paragraphs[0]
                for i in range(len(first_paragraph.runs)):
                    first_paragraph._p.remove(first_paragraph.runs[0]._r)
            else:
                first_paragraph = document.add_paragraph()

            paragraph_format = first_paragraph.paragraph_format
            paragraph_format.left_indent = Inches(0)
            paragraph_format.right_indent = Inches(0)
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)
            paragraph_format.line_spacing = 1.0

            run = first_paragraph.add_run()
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

    def replace_existing_image(self, docx_path: str, target_image_filename: str, new_image_path: Union[str, Path]) -> bool:
        """
        Replace an existing image in the DOCX file by modifying its underlying ZIP archive.
        Utiliza um diretório temporário que é automaticamente removido.
        
        Args:
            docx_path: Path to the DOCX file.
            target_image_filename: The filename of the image to be replaced (e.g., "image1.png").
            new_image_path: Path to the new image file.
            
        Returns:
            bool: True if the image was replaced successfully.
        """
        try:
            # Cria um diretório temporário que será automaticamente excluído
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extrai o conteúdo do DOCX para o diretório temporário
                with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                target_image_path = os.path.join(temp_dir, 'word', 'media', target_image_filename)
                if not os.path.exists(target_image_path):
                    raise FileNotFoundError(f"Target image '{target_image_filename}' not found in the DOCX file.")
                
                # Substitui a imagem antiga pela nova
                shutil.copy(str(new_image_path), target_image_path)
                
                # Reempacota o DOCX com os conteúdos modificados (sobrescrevendo o arquivo original)
                with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as new_docx:
                    for foldername, subfolders, filenames in os.walk(temp_dir):
                        for filename in filenames:
                            file_path = os.path.join(foldername, filename)
                            arcname = os.path.relpath(file_path, temp_dir)
                            new_docx.write(file_path, arcname)
                
                self.logger.info(f"Image replaced successfully in {docx_path}")
            # O diretório temporário é removido automaticamente ao sair do bloco 'with'
            return True

        except Exception as e:
            self.logger.error(f"Error replacing image: {str(e)}")
            return False

    def generate_report(self, context: dict, output_path: str,
                        cover_image_path: Optional[Union[str, Path]] = None,
                        target_image_filename: Optional[str] = "image1.png",
                        graau_params: dict = {}) -> bool:
        """
        Generates the report using the template and provided context.

        Args:
            context: Dictionary with the template context.
            output_path: Path to save the output file.
            cover_image_path: Optional path to the cover image.
            replace_image: If True, replaces an existing image in the DOCX file.
            target_image_filename: The filename of the image in the DOCX to replace (required if replace_image is True).
            graau_params: Additional parameters.
            
        Returns:
            bool: True if the report was generated successfully.
        """
        try:
            doc = DocxTemplate(self.template_path)

            if cover_image_path:
                # Render the document and save it temporarily
                temp_output = "temp_report.docx"
                doc.render(context)
                doc.save(temp_output)

                if not target_image_filename:
                    raise ValueError("target_image_filename must be provided.")

                if not self.replace_existing_image(temp_output, target_image_filename, cover_image_path):
                    return False

                # Move the temporary file to the final output path
                shutil.move(temp_output, output_path)
            else:
                doc.render(context)
                doc.save(output_path)

            self.logger.info(f"Report generated successfully: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"Error generating report: {str(e)}")
            return False

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    # Example context data
    sharepoint_data = [
        {
            "unidades_fiscalizadas": "P. M. CIDADE",
            "n_processo_eTCE": "TC/XXXXXX/20XX",
            "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
            "exercicios": "20XX, 20YY",
            "VRF": "R$ 100.000,00"
        }
    ]
    
    generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")
    
    success = generator.generate_report(
        context=sharepoint_data[0],
        output_path="src/reports/report_example.docx",
        cover_image_path="src/cover_images/cover_page_2.jpg",
    )
