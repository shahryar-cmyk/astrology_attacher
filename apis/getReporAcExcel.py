from flask import Blueprint, request, jsonify
import win32com.client
import pythoncom
import os
import subprocess

get_report_ac_excel = Blueprint('getReportAcExcel', __name__)

@get_report_ac_excel.route('/get_report_ac_excel', methods=['POST'])
def get_report_ac_excel_route():
    try:
        # Parse request data
        data = request.get_json()
        file_name = data.get('fileName')
        macro_name = data.get('macroName')
        file_list_path = r'C:\El Camino que Creas\Generador de Informes\Log.txt'  # Use a raw string for file path

        if not file_name or not macro_name:
            return jsonify({'error': 'fileName and macroName are required'}), 400
        
        # Check if Another Process is Running on the System (optional)
        # Uncomment to kill Excel processes
        # try:
        #     subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"])
        # except Exception as e:
        #     print("Error killing Excel process:", e)

        pythoncom.CoInitialize()  # Initialize COM library

        # Initialize Excel application
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False  # Set to True for debugging

        try:
            # Open the Excel workbook
            wb = xl.Workbooks.Open(file_name)
            try:
                # Access the specific sheet by name
                sheet_name = 'CN y RS (o RL)'  # Replace with your sheet name
                sheet = wb.Sheets(sheet_name)

                # Modify data in the sheet
                sheet.Range("S26").Value = "Hello, Shahryar!"

                # Run the macro
                xl.Application.Run(macro_name)

                # Now get the data in the Log file
                with open(file_list_path, 'r') as file:
                    file_names = [line.strip() for line in file.readlines()]

                if not file_names:
                    return jsonify({"error": "No files to process"}), 400

                # Get the last file name
                last_file_name = file_names[-1]

                # Initialize Word application
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False  # Make Word invisible

                # Open the Word document
                doc = word.Documents.Open(last_file_name)

                # Define the PDF file path based on the last file name
                pdf_file = os.path.splitext(last_file_name)[0] + '.pdf'  # Change extension to .pdf

                # Save as PDF
                doc.SaveAs(pdf_file, FileFormat=17)  # 17 is the format for PDF

                # Close the document and Word application
                doc.Close()
                word.Quit()

                return jsonify({"message": "File Saved Successfully", "fileName": last_file_name, "pdf_file": pdf_file}), 200
            except Exception as e:
                print("Error accessing the sheet or running macro:", e)
                return jsonify({"error": str(e)}), 500
            finally:
                wb.Close(SaveChanges=True)  # Save changes after modifying data
        except Exception as e:
            print("Error opening workbook:", e)
            return jsonify({"error": str(e)}), 500
        finally:
            xl.Quit()  # Quit the Excel application
        pythoncom.CoUninitialize()  # Uninitialize COM library
    except Exception as e:
        print("Error parsing request data:", e)
        return jsonify({"error": str(e)}), 400
