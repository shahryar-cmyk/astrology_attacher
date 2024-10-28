from flask import Blueprint, request, jsonify
import win32com.client
import pythoncom
import os

get_converted_pdf = Blueprint('get_converted_pdf', __name__)

@get_converted_pdf.route('/get_converted_pdf', methods=['POST'])
def get_converted_pdf_route():
    try:
        # Initialize COM library for multi-threaded environments
        pythoncom.CoInitialize()

        # Parse request data
        data = request.get_json()
        file_name = data.get('fileName')

        if not file_name:
            return jsonify({'error': 'fileName is required'}), 400

        # Ensure the file exists
        if not os.path.isfile(file_name):
            return jsonify({'error': f"File '{file_name}' not found"}), 404

        # Word application processing
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False  # Make Word invisible

        # Open the Word document
        doc = word.Documents.Open(file_name)

        # Define the PDF file path
        pdf_file = os.path.splitext(file_name)[0] + '.pdf'  # Change extension to .pdf

        # Save as PDF
        doc.SaveAs(pdf_file, FileFormat=17)  # 17 is the format for PDF

        # Close the document and Word application
        doc.Close()
        word.Quit()
        
        # Release COM library resources
        pythoncom.CoUninitialize()

        # Assuming successful operation, return success with PDF file path
        return jsonify({'message': 'PDF conversion successful', 'pdfFilePath': pdf_file}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500
