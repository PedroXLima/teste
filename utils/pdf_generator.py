"""
PDF generator — uses WeasyPrint to convert the institutional HTML to a PDF.
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)


class PdfGenerator:
    def __init__(self, base_url: Optional[str] = None):
        self.base_url = base_url or f'file://{os.path.abspath(os.path.dirname(__file__))}/'

    def generate(self, html_content: str, output_path: str) -> str:
        """
        Render *html_content* to a PDF file at *output_path*.
        Returns *output_path* on success.
        """
        try:
            from weasyprint import HTML, CSS
            from weasyprint.text.fonts import FontConfiguration

            font_config = FontConfiguration()

            # Optional extra CSS to ensure Google Fonts work even without internet
            fallback_css = CSS(
                string="""
                @font-face {
                    font-family: 'Montserrat';
                    src: local('Montserrat'), local('DejaVu Sans'), local('Liberation Sans');
                }
                @font-face {
                    font-family: 'Inter';
                    src: local('Inter'), local('DejaVu Sans'), local('Liberation Sans');
                }
                """,
                font_config=font_config,
            )

            doc = HTML(
                string=html_content,
                base_url=self.base_url,
            )

            doc.write_pdf(
                output_path,
                stylesheets=[fallback_css],
                font_config=font_config,
                presentational_hints=True,
            )

        except ImportError:
            raise RuntimeError(
                'WeasyPrint is not installed. Run: pip install weasyprint'
            )
        except Exception as exc:
            logger.exception('PDF generation failed: %s', exc)
            raise

        return output_path
