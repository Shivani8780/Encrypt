from flask import Flask, request, send_file, render_template
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
import io

app = Flask(__name__)

# Function to create watermark PDF in memory
def create_watermark(text, position):
    # Using landscape A4 page size (841.890 x 595.276)
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=landscape(A4))
    width, height = landscape(A4)
    
    # Set watermark text properties
    c.setFont("Helvetica-Bold", 14)  # Font size 14
    c.setFillColorRGB(0, 0, 1, alpha=0.5)  # Blue color with transparency

    # Positioning for watermark
    if position == 'top':
        x = width / 2 - (len(text) * 2.5)  # Center horizontally
        y = height - 15  # Very top of the page, slightly down (15 units)
    elif position == 'bottom':
        x = width / 2 - (len(text) * 2.5)  # Center horizontally
        y = 15  # Very bottom of the page, slightly up (15 units)

    # Draw watermark text
    c.drawString(x, y, text)
    c.save()
    
    # Return in-memory file
    packet.seek(0)
    return packet

# Function to apply watermark to PDF (using in-memory handling)
def add_watermark_to_pdf(input_pdf, watermark_pdf):
    reader = PdfReader(input_pdf)
    watermark = PdfReader(watermark_pdf)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        watermark_page = watermark.pages[0]
        
        # Merge watermark on top of the original page
        page.merge_page(watermark_page)
        writer.add_page(page)

    # Return in-memory PDF
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    output_pdf.seek(0)
    return output_pdf

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/encrypt', methods=['POST'])
def encrypt_pdf():
    if 'pdf_file' not in request.files or not request.form['password'] or not request.form['watermark_name'] or not request.form['custom_pdf_name']:
        return "Missing file, password, watermark name, or custom PDF name", 400

    pdf_file = request.files['pdf_file']
    password = request.form['password']
    watermark_name = request.form['watermark_name']
    custom_pdf_name = request.form['custom_pdf_name']

    if pdf_file.filename == '':
        return "No file selected", 400

    # Read the uploaded PDF into memory
    input_pdf = io.BytesIO(pdf_file.read())

    # Create custom watermarks in memory
    watermark_top = create_watermark(f"{watermark_name}", 'top')
    watermark_bottom = create_watermark(f"{watermark_name} ", 'bottom')

    try:
        # Apply watermarks
        watermarked_pdf = add_watermark_to_pdf(input_pdf, watermark_top)
        final_pdf = add_watermark_to_pdf(watermarked_pdf, watermark_bottom)

        # Encrypt the final PDF
        writer = PdfWriter()
        reader = PdfReader(final_pdf)
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)

        # Write the encrypted PDF to a BytesIO object (in-memory)
        output_pdf = io.BytesIO()
        writer.write(output_pdf)
        output_pdf.seek(0)

        # Send the encrypted PDF for download with the custom name
        return send_file(output_pdf, as_attachment=True, download_name=f"{custom_pdf_name}.pdf")

    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
