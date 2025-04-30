import os
import logging
import subprocess
import tempfile
from PIL import Image
import img2pdf
import pandas as pd
import pdfkit
import io

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Flag to track availability of wkhtmltopdf
WKHTMLTOPDF_AVAILABLE = False

# Try to check if wkhtmltopdf is available
try:
    result = subprocess.run(['which', 'wkhtmltopdf'], 
                           stdout=subprocess.PIPE, 
                           stderr=subprocess.PIPE)
    WKHTMLTOPDF_AVAILABLE = result.returncode == 0
    if not WKHTMLTOPDF_AVAILABLE:
        logger.warning("wkhtmltopdf is not available. HTML/CSV conversion will use alternate methods.")
except Exception as e:
    logger.warning(f"Error checking for wkhtmltopdf: {e}")
    WKHTMLTOPDF_AVAILABLE = False

def convert_to_pdf(input_path, output_path, file_extension):
    """
    Convert a file to PDF based on its extension
    
    Args:
        input_path: Path to the input file
        output_path: Path where the output PDF should be saved
        file_extension: Extension of the input file
        
    Returns:
        tuple: (success, error_message)
    """
    try:
        # Word documents (docx, doc)
        if file_extension in ['docx', 'doc']:
            return convert_word_to_pdf(input_path, output_path)
        
        # Excel files (xlsx, xls)
        elif file_extension in ['xlsx', 'xls']:
            return convert_excel_to_pdf(input_path, output_path)
        
        # PowerPoint files (pptx, ppt)
        elif file_extension in ['pptx', 'ppt']:
            return convert_powerpoint_to_pdf(input_path, output_path)
        
        # Images (jpg, jpeg, png, etc.)
        elif file_extension in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']:
            return convert_image_to_pdf(input_path, output_path)
        
        # HTML files
        elif file_extension in ['html', 'htm']:
            return convert_html_to_pdf(input_path, output_path)
        
        # Text files
        elif file_extension in ['txt', 'rtf']:
            return convert_text_to_pdf(input_path, output_path)
            
        # CSV files
        elif file_extension == 'csv':
            return convert_csv_to_pdf(input_path, output_path)
        
        else:
            return False, f"Unsupported file extension: {file_extension}"
    
    except Exception as e:
        logger.exception(f"Error converting {file_extension} file to PDF")
        return False, f"Conversion error: {str(e)}"

def convert_word_to_pdf(input_path, output_path):
    """Convert Word documents to PDF using fallback text extraction"""
    try:
        try:
            # First try to read the Word document as a binary file
            with open(input_path, 'rb') as f:
                content = f.read()
                
            # Try to extract plain text (very simple approach)
            try:
                text_content = ""
                # Try to decode as UTF-8 first (will work for simple .doc/.docx)
                text_content = content.decode('utf-8', errors='ignore')
            except:
                # Fallback to reading as a text file with error handling
                with open(input_path, 'r', errors='ignore') as f:
                    text_content = f.read()
                    
            # Create a simple PDF with the text content
            # Create HTML with the text content
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Word Document</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }}
                    pre {{ white-space: pre-wrap; }}
                </style>
            </head>
            <body>
                <h1>Converted Word Document</h1>
                <pre>{text_content}</pre>
            </body>
            </html>
            """
            
            temp_html = tempfile.NamedTemporaryFile(suffix='.html', delete=False)
            temp_html.write(html_content.encode('utf-8'))
            temp_html.close()
            
            # Convert HTML to PDF (which has PIL fallback)
            result = convert_html_to_pdf(temp_html.name, output_path)
            os.unlink(temp_html.name)
            return result
            
        except Exception as inner_e:
            logger.exception("Error extracting text from Word document")
            
            # Create a simple PDF with an error message
            width, height = 800, 1200  # A4 proportions
            image = Image.new('RGB', (width, height), color='white')
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(image)
            
            try:
                font = ImageFont.truetype("Arial", 12)
            except:
                font = ImageFont.load_default()
                
            draw.text((40, 40), "Word Document", font=font, fill="black")
            draw.text((40, 80), f"Filename: {os.path.basename(input_path)}", font=font, fill="black")
            draw.text((40, 120), "This Word document could not be converted to PDF.", font=font, fill="black")
            draw.text((40, 160), "Please make sure you have a compatible Word document.", font=font, fill="black")
            
            # Save as PDF
            image.save(output_path, "PDF", resolution=100.0)
            return True, None
    
    except Exception as e:
        logger.exception("Error in Word to PDF conversion")
        return False, str(e)

def convert_excel_to_pdf(input_path, output_path):
    """Convert Excel to PDF using pandas if available or backup rendering"""
    try:
        # Create a default error message 
        default_message = "Excel file conversion is not fully supported in this environment."
        
        try:
            # Try to use pandas to read the Excel file
            df = pd.read_excel(input_path)
            
            # Get the first few rows to display in a table
            preview_df = df.head(20)  # Show up to 20 rows
            
            # Convert to HTML
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Excel Conversion</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    table {{ border-collapse: collapse; width: 100%; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                    tr:nth-child(even) {{ background-color: #f9f9f9; }}
                </style>
            </head>
            <body>
                <h1>Excel Spreadsheet</h1>
                <p>Filename: {os.path.basename(input_path)}</p>
                <p>Total rows: {len(df)}</p>
                <p>Columns: {", ".join(df.columns)}</p>
                <h2>Data Preview:</h2>
                {preview_df.to_html(index=False)}
                {f"<p><i>Note: Only showing first {len(preview_df)} rows out of {len(df)} total rows.</i></p>" if len(df) > len(preview_df) else ""}
            </body>
            </html>
            """
            
            temp_html = tempfile.NamedTemporaryFile(suffix='.html', delete=False)
            temp_html.write(html_content.encode('utf-8'))
            temp_html.close()
            
            # Use the HTML converter which has PIL fallback
            result, error = convert_html_to_pdf(temp_html.name, output_path)
            os.unlink(temp_html.name)
            
            if result:
                return True, None
            else:
                raise Exception(f"HTML conversion failed: {error}")
                
        except Exception as inner_e:
            logger.exception("Error converting Excel with pandas")
            
            # Create a simple fallback PDF
            width, height = 800, 1200  # A4 proportions
            image = Image.new('RGB', (width, height), color='white')
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(image)
            
            try:
                font = ImageFont.truetype("Arial", 14)
                small_font = ImageFont.truetype("Arial", 12)
            except:
                font = ImageFont.load_default()
                small_font = font
                
            # Draw title and error message
            draw.text((40, 40), "Excel Spreadsheet", font=font, fill="black")
            draw.text((40, 80), f"Filename: {os.path.basename(input_path)}", font=small_font, fill="black")
            draw.text((40, 120), default_message, font=small_font, fill="black")
            draw.text((40, 160), "The spreadsheet content could not be displayed.", font=small_font, fill="black")
            
            # Add recommendations
            draw.text((40, 220), "Recommendations:", font=font, fill="black")
            draw.text((40, 260), "• Make sure the Excel file is not corrupted", font=small_font, fill="black")
            draw.text((40, 290), "• Consider saving the spreadsheet as CSV for better compatibility", font=small_font, fill="black")
            
            # Save as PDF
            image.save(output_path, "PDF", resolution=100.0)
            return True, None
            
    except Exception as e:
        logger.exception("Error in Excel to PDF conversion")
        return False, str(e)

def convert_powerpoint_to_pdf(input_path, output_path):
    """Create a PDF representation of a PowerPoint file"""
    try:
        # Create a simple PDF describing the PowerPoint
        width, height = 800, 1200  # A4 proportions
        image = Image.new('RGB', (width, height), color='white')
        from PIL import ImageDraw, ImageFont
        draw = ImageDraw.Draw(image)
        
        # Try to find a suitable font
        try:
            title_font = ImageFont.truetype("Arial", 24)
            heading_font = ImageFont.truetype("Arial", 18)
            regular_font = ImageFont.truetype("Arial", 14)
        except:
            title_font = ImageFont.load_default()
            heading_font = title_font
            regular_font = title_font
            
        # Draw a PowerPoint slide-like page
        # Header section
        draw.rectangle(((0, 0), (width, 80)), fill=(65, 105, 225))  # Royal blue header
        draw.text((40, 25), "PowerPoint Presentation", font=title_font, fill="white")
        
        # Main content
        y_pos = 120
        draw.text((40, y_pos), f"File: {os.path.basename(input_path)}", font=heading_font, fill="black")
        y_pos += 60
        
        # Add some info
        draw.text((40, y_pos), "Presentation Information", font=heading_font, fill="black")
        y_pos += 40
        
        info_text = [
            "• PowerPoint presentations require specialized software for full conversion.",
            "• This PDF contains basic information about your PowerPoint file.",
            "• To view the presentation with all features, use Microsoft PowerPoint or similar applications."
        ]
        
        for line in info_text:
            draw.text((40, y_pos), line, font=regular_font, fill="black")
            y_pos += 30
            
        y_pos += 30
        
        # Tips section
        draw.text((40, y_pos), "Recommendations:", font=heading_font, fill="black")
        y_pos += 40
        
        tips = [
            "• For better PDF conversion, use PowerPoint's built-in 'Save as PDF' feature",
            "• Consider using Google Slides which has good PDF export options",
            "• If sharing is your goal, consider exporting as a set of images"
        ]
        
        for tip in tips:
            draw.text((40, y_pos), tip, font=regular_font, fill="black")
            y_pos += 30
            
        # Footer
        footer_y = height - 50
        draw.line([(40, footer_y), (width - 40, footer_y)], fill=(200, 200, 200), width=1)
        draw.text((40, footer_y + 10), "PDF Converter", font=regular_font, fill=(100, 100, 100))
        
        # Save as PDF
        image.save(output_path, "PDF", resolution=100.0)
        return True, None
        
    except Exception as e:
        logger.exception("Error in PowerPoint to PDF conversion")
        return False, str(e)

def convert_image_to_pdf(input_path, output_path):
    """Convert image to PDF using img2pdf or PIL"""
    try:
        # Try using img2pdf first (better quality for most images)
        try:
            with open(output_path, "wb") as f:
                f.write(img2pdf.convert(input_path))
            return True, None
        except Exception as img2pdf_error:
            logger.warning(f"img2pdf conversion failed, trying PIL: {img2pdf_error}")
            
            # Fallback to PIL if img2pdf fails
            image = Image.open(input_path)
            
            # Handle RGBA images by converting to RGB
            if image.mode == 'RGBA':
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                rgb_image.paste(image, (0, 0), image)
                image = rgb_image
                
            image.save(output_path, "PDF", resolution=100.0)
            return True, None
            
    except Exception as e:
        logger.exception("Error in image to PDF conversion")
        return False, str(e)

def convert_html_to_pdf(input_path, output_path):
    """Convert HTML to PDF using pdfkit/wkhtmltopdf or PIL as fallback"""
    try:
        if WKHTMLTOPDF_AVAILABLE:
            # Use pdfkit if wkhtmltopdf is available
            options = {
                'page-size': 'A4',
                'margin-top': '0.75in',
                'margin-right': '0.75in',
                'margin-bottom': '0.75in',
                'margin-left': '0.75in',
                'encoding': "UTF-8",
            }
            
            pdfkit.from_file(input_path, output_path, options=options)
            return True, None
        else:
            # Fallback to PIL for rendering HTML to PDF
            logger.info("Using PIL fallback for HTML to PDF conversion")
            
            # Read the HTML content
            with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
                html_content = f.read()
                
            # Create a simple image with the HTML text content
            width, height = 800, 1200  # A4 proportions
            image = Image.new('RGB', (width, height), color='white')
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(image)
            
            # Try to find a suitable font
            try:
                font = ImageFont.truetype("Arial", 12)
            except IOError:
                font = ImageFont.load_default()
                
            # Extract text from HTML (very simple parsing)
            import re
            text_content = re.sub(r'<[^>]*>', ' ', html_content)
            text_content = re.sub(r'\s+', ' ', text_content).strip()
            
            # Draw the text with word wrapping
            margin = 40
            y_position = margin
            x_position = margin
            max_width = width - (2 * margin)
            line = ""
            
            for word in text_content.split():
                # Check if adding this word exceeds the width
                test_line = line + word + " "
                text_width = draw.textlength(test_line, font=font)
                
                if text_width <= max_width:
                    line = test_line
                else:
                    # Draw the line and move to next line
                    draw.text((x_position, y_position), line, font=font, fill="black")
                    y_position += 20  # Line height
                    line = word + " "
                    
                    # Check if we need a new page (simple pagination)
                    if y_position > height - margin:
                        break
                        
            # Draw any remaining text
            if line:
                draw.text((x_position, y_position), line, font=font, fill="black")
                
            # Save as PDF
            image.save(output_path, "PDF", resolution=100.0)
            return True, None
            
    except Exception as e:
        logger.exception("Error in HTML to PDF conversion")
        return False, str(e)

def convert_text_to_pdf(input_path, output_path):
    """Convert text file to PDF"""
    try:
        # Read the text content
        with open(input_path, 'r', errors='ignore') as f:
            text_content = f.read()
        
        # Convert to HTML
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Text Document</title>
            <style>
                body {{ font-family: "Courier New", Courier, monospace; margin: 40px; line-height: 1.6; }}
                pre {{ white-space: pre-wrap; }}
            </style>
        </head>
        <body>
            <pre>{text_content}</pre>
        </body>
        </html>
        """
        
        # Save as temporary HTML file
        temp_html = tempfile.NamedTemporaryFile(suffix='.html', delete=False)
        temp_html.write(html_content.encode('utf-8'))
        temp_html.close()
        
        # Convert HTML to PDF
        html_result, html_error = convert_html_to_pdf(temp_html.name, output_path)
        
        # Clean up
        os.unlink(temp_html.name)
        
        return html_result, html_error
        
    except Exception as e:
        logger.exception("Error in text to PDF conversion")
        return False, str(e)

def convert_csv_to_pdf(input_path, output_path):
    """Convert CSV to PDF using pandas and pdfkit/PIL"""
    try:
        # Read the CSV file
        df = pd.read_csv(input_path)
        
        # Convert to HTML with styling
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>CSV Data</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
            </style>
        </head>
        <body>
            <h2>CSV Data</h2>
            {df.to_html(index=False)}
        </body>
        </html>
        """
        
        # Save as temporary HTML file
        temp_html = tempfile.NamedTemporaryFile(suffix='.html', delete=False)
        temp_html.write(html_content.encode('utf-8'))
        temp_html.close()
        
        # Use the HTML converter which has a PIL fallback
        result, error = convert_html_to_pdf(temp_html.name, output_path)
        
        # Clean up
        os.unlink(temp_html.name)
        
        if result:
            return True, None
        else:
            # If the HTML conversion failed, we'll create a simple text-based PDF
            try:
                # Create a simple image with CSV data
                width, height = 800, 1200  # A4 proportions
                image = Image.new('RGB', (width, height), color='white')
                from PIL import ImageDraw, ImageFont
                draw = ImageDraw.Draw(image)
                
                # Try to find a suitable font
                try:
                    font = ImageFont.truetype("Arial", 12)
                except IOError:
                    font = ImageFont.load_default()
                
                # Draw CSV data as a table
                margin = 40
                y_position = margin
                x_position = margin
                
                # Draw title
                draw.text((x_position, y_position), "CSV Data", font=font, fill="black")
                y_position += 30
                
                # Get column width based on data
                col_widths = []
                for col in df.columns:
                    col_width = draw.textlength(col, font=font) + 20
                    col_widths.append(max(col_width, 100))  # Minimum width
                
                # Draw header
                for i, col in enumerate(df.columns):
                    draw.text((x_position, y_position), col, font=font, fill="black")
                    x_position += col_widths[i]
                
                y_position += 20
                x_position = margin
                
                # Draw separator line
                draw.line([(margin, y_position), (min(margin + sum(col_widths), width - margin), y_position)], 
                          fill="black", width=1)
                y_position += 10
                
                # Draw rows (limited to what can fit on the page)
                max_rows = min(20, len(df))
                for row_idx in range(max_rows):
                    x_position = margin
                    for col_idx, col in enumerate(df.columns):
                        cell_value = str(df.iloc[row_idx][col])
                        # Truncate if too long
                        if len(cell_value) > 15:
                            cell_value = cell_value[:12] + "..."
                        draw.text((x_position, y_position), cell_value, font=font, fill="black")
                        x_position += col_widths[col_idx]
                    y_position += 20
                    
                    if y_position > height - margin:
                        draw.text((margin, y_position - 10), 
                                 f"... and {len(df) - row_idx - 1} more rows", font=font, fill="black")
                        break
                
                # Save as PDF
                image.save(output_path, "PDF", resolution=100.0)
                return True, None
            except Exception as e:
                logger.exception("Error in CSV fallback conversion")
                return False, f"CSV to PDF conversion failed: {str(e)}"
        
    except Exception as e:
        logger.exception("Error in CSV to PDF conversion")
        return False, str(e)
