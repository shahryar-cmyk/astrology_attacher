from flask import Blueprint, request, jsonify
import win32com.client
import pythoncom
import os
import subprocess

get_report_ac_excel = Blueprint('getReportAcExcel', __name__)

listofMacros = [
    {
        "macroId": 1,
        "macroName": "GI1CasaCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
        "IsCompleted": False
    
    },
    {
        "macroId": 2,
        "macroName": "GI1CasaRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 3,
        "macroName": "GI1CasaRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 4,
        "macroName": "GI1CasaSecretaCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 5,
        "macroName": "GI1CasaSecretaRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 6,
        "macroName": "GI1CasaSecretaRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 7,
        "macroName": "GI1DeCCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 8,
        "macroName": "GI1DeCRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 9,
        "macroName": "GI1DeCRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 10,
        "macroName": "GI1DeSCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
    },
    {
        "macroId": 11,
        "macroName": "GI1DeSRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 12,
        "macroName": "GI1DeSRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 13,
        "macroName": "GI1DSCCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 14,
        "macroName": "GI1DSCRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 15,
        "macroName": "GI1DSCRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 16,
        "macroName": "GI1SignoCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 17,
        "macroName": "GI1SignoRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 18,
        "macroName": "GI1SignoRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 19,
        "macroName": "GI2CasaCN",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
    {
        "macroId": 20,
        "macroName": "GI2CasaRL",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
    },
    {
        "macroId": 21,
        "macroName": "GI2CasaRS",
        "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm",
    },
# Upto ID 72
{
    "macroId": 22,
    "macroName": "GI2CasaSecretaCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 23,
    "macroName": "GI2CasaSecretaRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 24,
    "macroName": "GI2CasaSecretaRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 25,
    "macroName": "GI2DeCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 26,
    "macroName": "GI2DeCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 27,
    "macroName": "GI2DeCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 28,
    "macroName": "GI2DeSCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 29,
    "macroName": "GI2DeSRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 30,
    "macroName": "GI2DeSRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 31,
    "macroName": "GI2DSCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 32,
    "macroName": "GI2DSCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 33,
    "macroName": "GI2DSCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 34,
    "macroName": "GI2SignoCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 35,
    "macroName": "GI2SignoRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 36,
    "macroName": "GI2SignoRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 37,
    "macroName": "GI3CasaCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 38,
    "macroName": "GI3CasaRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 39,
    "macroName": "GI3CasaRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 40,
    "macroName": "GI3CasaSecretaCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 41,
    "macroName": "GI3CasaSecretaRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 42,
    "macroName": "GI3CasaSecretaRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 43,
    "macroName": "GI3DeCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 44,
    "macroName": "GI3DeCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 45,
    "macroName": "GI3DeCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 46,
    "macroName": "GI3DeSCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 47,
    "macroName": "GI3DeSRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 48,
    "macroName": "GI3DeSRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 49,
    "macroName": "GI3DSCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 50,
    "macroName": "GI3DSCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 51,
    "macroName": "GI3DSCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 52,
    "macroName": "GI3SignoCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 53,
    "macroName": "GI3SignoRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 54,
    "macroName": "GI3SignoRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 55,
    "macroName": "GI4CasaCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 56,
    "macroName": "GI4CasaRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 57,
    "macroName": "GI4CasaRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 58,
    "macroName": "GI4CasaSecretaCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 59,
    "macroName": "GI4CasaSecretaRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 60,
    "macroName": "GI4CasaSecretaRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 61,
    "macroName": "GI4DeCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 62,
    "macroName": "GI4DeCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 63,
    "macroName": "GI4DeCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 64,
    "macroName": "GI4DeSCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 65,
    "macroName": "GI4DeSRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 66,
    "macroName": "GI4DeSRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 67,
    "macroName": "GI4DSCCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 68,
    "macroName": "GI4DSCRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 69,
    "macroName": "GI4DSCRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 70,
    "macroName": "GI4SignoCN",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 71,
    "macroName": "GI4SignoRL",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
},
{
    "macroId": 72,
    "macroName": "GI4SignoRS",
    "macroNameInWordPress": "Generador de Informes_20240708_154017_929359.xlsm"
}
]

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
        
        # Check if Another Process is Running on the System
        try:
            subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"])
        except Exception as e:
            print("Error killing Excel process:", e)
        
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

                for file_name in file_names:
                    if file_name != last_file_name:
                        try:
                            os.remove(file_name)
                            print(f"Deleted {file_name}")
                        except FileNotFoundError:
                            print(f"File {file_name} not found.")
                        except Exception as e:
                            print(f"Error deleting file {file_name}: {e}")

                return jsonify({"message": "File Saved Successfully", "fileName": file_name}), 200
            except Exception as e:
                print("Error accessing the sheet:", e)
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