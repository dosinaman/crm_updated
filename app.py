from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
@app.route('/', methods=['GET', 'POST'])
def report_form():
    if request.method == 'POST':
        # Get form data
        ar_no = request.form['ar_no']
        product = request.form['product']
        analysed_on = request.form['analysed_on']
        purchaser = request.form['purchaser']
        batch_no = request.form['batch_no']
        po_no = request.form['po_no']
        date = request.form['date']
        qty_ordered = request.form['qty_ordered']
        qty_dispatched = request.form['qty_dispatched']
        test_results = request.form['test_results']  # Get test results in JSON format
        
        # Create PDF
        pdf_io = create_pdf(ar_no, product, analysed_on, purchaser, batch_no, po_no, date, qty_ordered, qty_dispatched, test_results)
        return send_file(pdf_io, as_attachment=True, download_name='analysis_report.pdf', mimetype='application/pdf')

    return render_template('report_form.html')

def create_pdf(ar_no, product, analysed_on, purchaser, batch_no, po_no, date, qty_ordered, qty_dispatched, test_results):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=letter)

    # Company Name
    c.setFont("Helvetica-Bold", 36)
    c.setFillColorRGB(1, 0, 0)  # Set color to red
    c.drawString(200, 750, "K.G. INDUSTRIES")
    
    # Company Description
    c.setFont("Helvetica-Bold", 11)
    c.setFillColorRGB(0, 0, 0)  # Set color back to black
    c.drawString(20, 725, "Mfg. : Magnesium Metal Powder, Turning & Specialized in Non Ferro Metal Powder and Chemical Suppliers")
    
    # Company Address
    c.setFont("Helvetica", 11)
    c.drawString(130, 700, "202 Ankur Palace, Scheme No. 54, Vijay Nagar, A.B. Road, INDORE- 452010(M.P.)")
    
    # Reference Number and Date
    c.setFont("Helvetica-Bold", 11)
    c.drawString(20, 675, "Ref. No.:")
    c.drawString(500, 675, "Date: " + date)

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 640, "CERTIFICATE OF ANALYSIS")

    # Draw a table for report details
    c.setFont("Helvetica-Bold", 12)

    # Data for the table
    data = [
        ["TITLE", "CERTIFICATE OF ANALYSIS", "A.R. No.", ar_no],
        ["PRODUCT", product, "ANALYSED ON", analysed_on],
        ["PURCHASER", purchaser, "BATCH No.", batch_no],
        ["P.O. No.", po_no, "Date", date],
        ["QTY ORDERED:", f"{qty_ordered} KG", "QTY. DISPATCHED", f"{qty_dispatched} KG"]
    ]

    # Define the position and size of the table
    x_start = 50
    y_start = 600
    cell_width = 150
    cell_height = 40

    # Draw the table
    for row_num, row in enumerate(data):
        for col_num, cell in enumerate(row):
            x = x_start + col_num * cell_width
            y = y_start - row_num * cell_height
            c.drawString(x + 5, y - 25, cell)  # 5 is for padding from the cell border
            c.rect(x, y - cell_height, cell_width, cell_height)  # Draw cell border

    # Add a line separator for test results
    c.line(100, y - 10, 500, y - 10)

    
    # # Draw table headers
    # c.drawString(100, 620, "A.R. No.")
    # c.drawString(250, 620, "PRODUCT")
    # c.drawString(400, 620, "ANALYSED ON")
    
    # # Fill in the report details
    # c.setFont("Helvetica", 12)
    # c.drawString(100, 600, ar_no)
    # c.drawString(250, 600, product)
    # c.drawString(400, 600, analysed_on)

    # c.setFont("Helvetica-Bold", 12)
    # c.drawString(100, 580, "PURCHASER")
    # c.drawString(250, 580, "BATCH No.")
    # c.drawString(400, 580, "P.O. No.")
    
    # c.setFont("Helvetica", 12)
    # c.drawString(100, 560, purchaser)
    # c.drawString(250, 560, batch_no)
    # c.drawString(400, 560, po_no)

    # c.setFont("Helvetica-Bold", 12)
    # c.drawString(100, 540, "QTY ORDERED:")
    # c.drawString(250, 540, "QTY DISPATCHED:")
    
    # c.setFont("Helvetica", 12)
    # c.drawString(100, 520, f"{qty_ordered} KG")
    # c.drawString(250, 520, f"{qty_dispatched} KG")

    # # Add a line separator
    # c.line(100, 510, 500, 510)

    # Add table header for test results
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, 490, "S. NO.")
    c.drawString(150, 490, "TEST")
    c.drawString(300, 490, "SPECIFICATION")
    c.drawString(450, 490, "RESULTS")

    # Parse test results JSON
    import json
    tests = json.loads(test_results)

    y_position = 470
    for i, test in enumerate(tests, start=1):
        c.setFont("Helvetica", 12)
        c.drawString(100, y_position, str(i) + ".")
        c.drawString(150, y_position, test['test'])
        c.drawString(300, y_position, test['specification'])
        c.drawString(450, y_position, test['results'])
        y_position -= 20  # Adjust for next test entry

    c.save()
    pdf_io.seek(0)
    return pdf_io


if __name__ == '__main__':
    app.run(debug=True)
