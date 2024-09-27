from flask import Blueprint, jsonify, request
import subprocess
import re
import win32com.client
import pythoncom
import os
import shutil
from datetime import datetime,timedelta
import logging
import traceback
import swisseph as swe


calculate_sun_return = Blueprint('calculate_sun_return', __name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
# Mapeo de las abreviaturas de los signos del zodíaco a sus nombres completos
zodiac_signs = {
    'ar': 'Aries',
    'ta': 'Tauro',
    'ge': 'Géminis',
    'cn': 'Cáncer',
    'le': 'Leo',
    'vi': 'Virgo',
    'li': 'Libra',
    'sc': 'Escorpio',
    'sa': 'Sagitario',
    'cp': 'Capricornio',
    'aq': 'Acuario',
    'pi': 'Piscis'
    # Agrega otras abreviaturas si es necesario
}

@calculate_sun_return.route('/calculate_sun_return', methods=['POST'])
def run_excel_macro_changeData():
    try:
        subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"])
        close_excel_without_save()
    except Exception as e:
            print("Error killing Excel process:", e)
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Get the parameters from the request data and ensure they are integers
        birth_date_year = int(request.json.get('birth_date_year'))
        birth_date_month = int(request.json.get('birth_date_month'))
        birth_date_day = int(request.json.get('birth_date_day'))
        ut_hour = int(request.json.get('ut_hour'))
        ut_min = int(request.json.get('ut_min'))
        ut_sec = int(request.json.get('ut_sec'))
        lat_deg = request.json.get('lat_deg')
        lon_deg = request.json.get('lon_deg')
        # Moon Return, Solar Return or Natal return 
        report_type_data = request.json.get('reportType')
        person_name = request.json.get('personName')
        person_location = request.json.get('personLocation')
        person_birth_date_local = request.json.get('personBirthDateLocal')
        sun_return_date = request.json.get('sunReturnDate')
        gender_type = request.json.get('gender')

        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False  # Set to True if you want Excel to be visible

        # Construct the command with zero-padded values
        # For House Data From Cell D5 to D10
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
        # For Planets Data From Cell D11 to D21 Which Includes True Node
        command2 = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZS -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"
        # For Quirón Command From Cell D22sky
        quiron_planet = f"swetest -ps -xs2060 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Lilith Command From Cell D23
        lilith_planet = f"swetest -ps -xs1181 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Cerus Command 
        cerus_planet = f"swetest -ps -xs1 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Pallas Command
        pallas_planet = f"swetest -ps -xs2 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Juno Command
        juno_planet = f"swetest -ps -xs3 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Vesta Command
        vesta_planet = f"swetest -ps -xs4 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Eris Command
        eris_planet = f"swetest -ps -xs136199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # White Moon Command
        white_moon = f"swetest -pZ -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Quaoar Command
        quaoar_planet = f"swetest -ps -xs50000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Sedna Command
        sedna_planet = f"swetest -ps -xs90377 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Varuna Command
        varuna_planet = f"swetest -ps -xs20000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Nessus Command
        nessus_planet = f"swetest -ps -xs7066 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Waltemath Command
        waltemath_planet = f"swetest -pw -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hygeia Command
        hygeia_planet = f"swetest -ps -xs10 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Sylvia Command
        sylvia_planet = f"swetest -ps -xs87 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
#         Hektor	624	Hector
        hektor_planet = f"swetest -ps -xs624 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Europa	52	Europa
        europa_planet = f"swetest -ps -xs52 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Davida	511	Davida
        davida_planet = f"swetest -ps -xs511 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Interamnia	704	Interamnia
        interamnia_planet = f"swetest -ps -xs704 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Camilla	107	Camilla
        camilla_planet = f"swetest -ps -xs107 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Cybele	65	Cybele
        cybele_planet = f"swetest -ps -xs65 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Sol Negro	h22	Black Sun

# Antivertex		Anti-Vertex

# Nodo Sur Real		True South Node
# Sol Negro Real		True Black Sun
# Lilith 2		Lilith 2
# Waldemath Priapus		Waldemath Priapus
# Sol Blanco		White Sun
# Chariklo	10199	Chariklo
        chariklo_planet = f"swetest -ps -xs10199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Iris	7	Iris
        iris_planet = f"swetest -ps -xs7 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Eunomia	15	Eunomia
        eunomia_planet = f"swetest -ps -xs15 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Euphrosyne	31	Euphrosyne
        euphrosyne_planet = f"swetest -ps -xs31 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Orcus	90482	Orcus
        # Orcus Command
        orcus_planet = f"swetest -ps -xs90482 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Pholus	5145	Pholus
        # Pholus Command
        pholus_planet = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hermíone Command
        hermione_planet = f"swetest -ps -xs121 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ixion Command
        ixion_planet = f"swetest -ps -xs28978 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Haumea Command
        haumea_planet = f"swetest -ps -xs136108 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Makemake Command
        makemake_planet = f"swetest -ps -xs136472 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Bamberga Command
        bamberga_planet = f"swetest -ps -xs324 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Patientia Command
        patientia_planet = f"swetest -ps -xs451 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Thisbe Command
        thisbe_planet = f"swetest -ps -xs88 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Herculina Command
        herculina_planet = f"swetest -ps -xs532 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Doris Command
        doris_planet = f"swetest -ps -xs48 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ursula Command
        ursula_planet = f"swetest -ps -xs375 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Eugenia Command
        eugenia_planet = f"swetest -ps -xs45 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Amphitrite Command
        amphitrite_planet = f"swetest -ps -xs29 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Diotima Command
        diotima_planet = f"swetest -ps -xs423 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Fortuna Command
        fortuna_planet = f"swetest -ps -xs19 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Egeria Command
        egeria_planet = f"swetest -ps -xs13 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Themis Command
        themis_planet = f"swetest -ps -xs24 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aurora Command
        aurora_planet = f"swetest -ps -xs94 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Alauda Command
        alauda_planet = f"swetest -ps -xs702 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aletheia Command
        aletheia_planet = f"swetest -ps -xs259 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Palma Command
        palma_planet = f"swetest -ps -xs372 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Nemesis Command
        nemesis_planet = f"swetest -ps -xs128 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Psyche Command
        psyche_planet = f"swetest -ps -xs16 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hebe Command
        hebe_planet = f"swetest -ps -xs6 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lachesis Command
        lachesis_planet = f"swetest -ps -xs120 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Daphne Command
        daphne_planet = f"swetest -ps -xs41 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Bertha Command
        bertha_planet = f"swetest -ps -xs154 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Freia Command
        freia_planet = f"swetest -ps -xs76 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Winchester Command
        winchester_planet = f"swetest -ps -xs747 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hilda Command
        hilda_planet = f"swetest -ps -xs153 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Pretoria Command
        pretoria_planet = f"swetest -ps -xs790 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Metis Command
        metis_planet = f"swetest -ps -xs9 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aegle Command
        aegle_planet = f"swetest -ps -xs96 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Kalliope Command
        kalliope_planet = f"swetest -ps -xs22 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Germania Command
        germania_planet = f"swetest -ps -xs241 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Prokne Command
        prokne_planet = f"swetest -ps -xs194 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Stereoskopia Command
        stereoskopia_planet = f"swetest -ps -xs566 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Agamemnon Command
        agamemnon_planet = f"swetest -ps -xs911 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Alexandra Command
        alexandra_planet = f"swetest -ps -xs54 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Siegena Command
        siegena_planet = f"swetest -ps -xs386 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Elpis Command
        elpis_planet = f"swetest -ps -xs59 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lilith Real Command
        # osc. Apogee Command in pa
        lilith_real_planet = f"swetest -pa -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Borasisi Command
        borasisi_planet = f"swetest -ps -xs66652 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lempo Command
        lempo_planet = f"swetest -ps -xs47171 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # 1998(26308) Command
        _1998_26308_planet = f"swetest -ps -xs26308 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ceto Command
        ceto_planet = f"swetest -ps -xs65489 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Teharonhiawako Command
        teharonhiawako_planet = f"swetest -ps -xs88611 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # 2000 OJ67 (134860) Command
        _2000_oj67_134860_planet = f"swetest -ps -xs134860 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Elektra Command
        elektra_planet = f"swetest -ps -xs130 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Typhon Command
        typhon_planet = f"swetest -ps -xs42355 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aspasia Command
        aspasia_planet = f"swetest -ps -xs409 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Chicago Command
        chicago_planet = f"swetest -ps -xs334 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Loreley Command
        loreley_planet = f"swetest -ps -xs165 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Gyptis Command
        gyptis_planet = f"swetest -ps -xs444 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Diomedes Command
        diomedes_planet = f"swetest -ps -xs1437 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Kreusa Command
        kreusa_planet = f"swetest -ps -xs488 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Juewa	139	Juewa
        juewa_planet = f"swetest -ps -xs139 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Eunike	185	Eunike
        eunike_planet = f"swetest -ps -xs185 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Ino	173	Ino
        ino_planet = f"swetest -ps -xs173 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Ismene	190	Ismene
        ismene_planet = f"swetest -ps -xs190 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Merapi	536	Merapi
        merapi_planet = f"swetest -ps -xs536 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"


        # Execute the command using subprocess
        # Planet Names
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        # Houses Names
        result2 = subprocess.run(command2, shell=True, check=True, capture_output=True, text=True)
        quiron_planet_result = subprocess.run(quiron_planet, shell=True, check=True, capture_output=True, text=True)
        lilith_planet_result = subprocess.run(lilith_planet, shell=True, check=True, capture_output=True, text=True)
        cerus_planet_result = subprocess.run(cerus_planet, shell=True, check=True, capture_output=True, text=True)
        pallas_planet_result = subprocess.run(pallas_planet, shell=True, check=True, capture_output=True, text=True)
        juno_planet_result = subprocess.run(juno_planet, shell=True, check=True, capture_output=True, text=True)
        vesta_planet_result = subprocess.run(vesta_planet, shell=True, check=True, capture_output=True, text=True)
        eris_planet_result = subprocess.run(eris_planet, shell=True, check=True, capture_output=True, text=True)
        white_moon_result = subprocess.run(white_moon, shell=True, check=True, capture_output=True, text=True)
        quaoar_planet_result = subprocess.run(quaoar_planet, shell=True, check=True, capture_output=True, text=True)
        sedna_planet_result = subprocess.run(sedna_planet, shell=True, check=True, capture_output=True, text=True)
        varuna_planet_result = subprocess.run(varuna_planet, shell=True, check=True, capture_output=True, text=True)
        nessus_planet_result = subprocess.run(nessus_planet, shell=True, check=True, capture_output=True, text=True)
        waltemath_planet_result = subprocess.run(waltemath_planet, shell=True, check=True, capture_output=True, text=True)
        hygeia_planet_result = subprocess.run(hygeia_planet, shell=True, check=True, capture_output=True, text=True)
        sylvia_planet_result = subprocess.run(sylvia_planet, shell=True, check=True, capture_output=True, text=True)
        hektor_planet_result = subprocess.run(hektor_planet, shell=True, check=True, capture_output=True, text=True)
        europa_planet_result = subprocess.run(europa_planet, shell=True, check=True, capture_output=True, text=True)
        davida_planet_result = subprocess.run(davida_planet, shell=True, check=True, capture_output=True, text=True)
        interamnia_planet_result = subprocess.run(interamnia_planet, shell=True, check=True, capture_output=True, text=True)
        camilla_planet_result = subprocess.run(camilla_planet, shell=True, check=True, capture_output=True, text=True)
        cybele_planet_result = subprocess.run(cybele_planet, shell=True, check=True, capture_output=True, text=True)
        chariklo_planet_result = subprocess.run(chariklo_planet, shell=True, check=True, capture_output=True, text=True)
        iris_planet_result = subprocess.run(iris_planet, shell=True, check=True, capture_output=True, text=True)
        eunomia_planet_result = subprocess.run(eunomia_planet, shell=True, check=True, capture_output=True, text=True)
        euphrosyne_planet_result = subprocess.run(euphrosyne_planet, shell=True, check=True, capture_output=True, text=True)
        orcus_planet_result = subprocess.run(orcus_planet, shell=True, check=True, capture_output=True, text=True)
        pholus_planet_result = subprocess.run(pholus_planet, shell=True, check=True, capture_output=True, text=True)
        hermione_planet_result = subprocess.run(hermione_planet, shell=True, check=True, capture_output=True, text=True)
        ixion_planet_result = subprocess.run(ixion_planet, shell=True, check=True, capture_output=True, text=True)
        haumea_planet_result = subprocess.run(haumea_planet, shell=True, check=True, capture_output=True, text=True)
        makemake_planet_result = subprocess.run(makemake_planet, shell=True, check=True, capture_output=True, text=True)
        bamberga_planet_result = subprocess.run(bamberga_planet, shell=True, check=True, capture_output=True, text=True)
        patientia_planet_result = subprocess.run(patientia_planet, shell=True, check=True, capture_output=True, text=True)
        thisbe_planet_result = subprocess.run(thisbe_planet, shell=True, check=True, capture_output=True, text=True)
        herculina_planet_result = subprocess.run(herculina_planet, shell=True, check=True, capture_output=True, text=True)
        doris_planet_result = subprocess.run(doris_planet, shell=True, check=True, capture_output=True, text=True)
        ursula_planet_result = subprocess.run(ursula_planet, shell=True, check=True, capture_output=True, text=True)
        eugenia_planet_result = subprocess.run(eugenia_planet, shell=True, check=True, capture_output=True, text=True)
        amphitrite_planet_result = subprocess.run(amphitrite_planet, shell=True, check=True, capture_output=True, text=True)
        diotima_planet_result = subprocess.run(diotima_planet, shell=True, check=True, capture_output=True, text=True)
        fortuna_planet_result = subprocess.run(fortuna_planet, shell=True, check=True, capture_output=True, text=True)
        egeria_planet_result = subprocess.run(egeria_planet, shell=True, check=True, capture_output=True, text=True)
        themis_planet_result = subprocess.run(themis_planet, shell=True, check=True, capture_output=True, text=True)
        aurora_planet_result = subprocess.run(aurora_planet, shell=True, check=True, capture_output=True, text=True)
        alauda_planet_result = subprocess.run(alauda_planet, shell=True, check=True, capture_output=True, text=True)
        aletheia_planet_result = subprocess.run(aletheia_planet, shell=True, check=True, capture_output=True, text=True)
        palma_planet_result = subprocess.run(palma_planet, shell=True, check=True, capture_output=True, text=True)
        nemesis_planet_result = subprocess.run(nemesis_planet, shell=True, check=True, capture_output=True, text=True)
        psyche_planet_result = subprocess.run(psyche_planet, shell=True, check=True, capture_output=True, text=True)
        hebe_planet_result =  subprocess.run(hebe_planet, shell=True, check=True, capture_output=True, text=True)
        lachesis_planet_result = subprocess.run(lachesis_planet, shell=True, check=True, capture_output=True, text=True)
        daphne_planet_result = subprocess.run(daphne_planet, shell=True, check=True, capture_output=True, text=True)
        bertha_planet_result = subprocess.run(bertha_planet, shell=True, check=True, capture_output=True, text=True)
        freia_planet_result = subprocess.run(freia_planet, shell=True, check=True, capture_output=True, text=True)
        winchester_planet_result = subprocess.run(winchester_planet, shell=True, check=True, capture_output=True, text=True)
        hilda_planet_result = subprocess.run(hilda_planet, shell=True, check=True, capture_output=True, text=True)
        pretoria_planet_result = subprocess.run(pretoria_planet, shell=True, check=True, capture_output=True, text=True)
        metis_planet_result = subprocess.run(metis_planet, shell=True, check=True, capture_output=True, text=True)
        aegle_planet_result = subprocess.run(aegle_planet, shell=True, check=True, capture_output=True, text=True)
        kalliope_planet_result = subprocess.run(kalliope_planet, shell=True, check=True, capture_output=True, text=True)
        germania_planet_result = subprocess.run(germania_planet, shell=True, check=True, capture_output=True, text=True)
        prokne_planet_result = subprocess.run(prokne_planet, shell=True, check=True, capture_output=True, text=True)
        stereoskopia_planet_result = subprocess.run(stereoskopia_planet, shell=True, check=True, capture_output=True, text=True)
        agamemnon_planet_result = subprocess.run(agamemnon_planet, shell=True, check=True, capture_output=True, text=True)
        alexandra_planet_result = subprocess.run(alexandra_planet, shell=True, check=True, capture_output=True, text=True)
        siegena_planet_result = subprocess.run(siegena_planet, shell=True, check=True, capture_output=True, text=True)
        elpis_planet_result = subprocess.run(elpis_planet, shell=True, check=True, capture_output=True, text=True)
        lilith_real_planet_result = subprocess.run(lilith_real_planet, shell=True, check=True, capture_output=True, text=True)
        borasisi_planet_result = subprocess.run(borasisi_planet, shell=True, check=True, capture_output=True, text=True)
        lempo_planet_result = subprocess.run(lempo_planet, shell=True, check=True, capture_output=True, text=True)
        _1998_26308_planet_result = subprocess.run(_1998_26308_planet, shell=True, check=True, capture_output=True, text=True)
        ceto_planet_result = subprocess.run(ceto_planet, shell=True, check=True, capture_output=True, text=True)
        teharonhiawako_planet_result = subprocess.run(teharonhiawako_planet, shell=True, check=True, capture_output=True, text=True)
        _2000_oj67_134860_planet_result = subprocess.run(_2000_oj67_134860_planet, shell=True, check=True, capture_output=True, text=True)
        elektra_planet_result = subprocess.run(elektra_planet, shell=True, check=True, capture_output=True, text=True)
        typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        chicago_planet_result = subprocess.run(chicago_planet, shell=True, check=True, capture_output=True, text=True)
        loreley_planet_result = subprocess.run(loreley_planet, shell=True, check=True, capture_output=True, text=True)
        diomedes_planet_result = subprocess.run(diomedes_planet, shell=True, check=True, capture_output=True, text=True)
        kreusa_planet_result = subprocess.run(kreusa_planet, shell=True, check=True, capture_output=True, text=True)
        gyptis_planet_result = subprocess.run(gyptis_planet, shell=True, check=True, capture_output=True, text=True)
        juewa_planet_result = subprocess.run(juewa_planet, shell=True, check=True, capture_output=True, text=True)
        eunike_planet_result = subprocess.run(eunike_planet, shell=True, check=True, capture_output=True, text=True)
        ino_planet_result = subprocess.run(ino_planet, shell=True, check=True, capture_output=True, text=True)
        ismene_planet_result = subprocess.run(ismene_planet, shell=True, check=True, capture_output=True, text=True)
        merapi_planet_result = subprocess.run(merapi_planet, shell=True, check=True, capture_output=True, text=True)

        astro_objects = [
    "Chiron",
    "Lilith",
    "Vertex",
    "Ceres",
    "Pallas",
    "Juno",
    "Vesta",
    "Eris",
    "Selena/White Moon",
    "Quaoar",
    "Sedna",
    "Varuna",
    "Nessus",
    "Waldemath",
    "Hygiea",
    "Sylvia",
    "Hektor",
    "Europa",
    "Davida",
    "Interamnia",
    "Camilla",
    "Cybele",
    "Black Sun",
    "Anti-Vertex",
    "True South Node",
    "True Black Sun",
    "Lilith 2",
    "Waldemath Priapus",
    "White Sun",
    "Chariklo",
    "Iris",
    "Eunomia",
    "Euphrosyne",
    "Orcus",
    "Pholus",
    "Hermione",
    "Ixion",
    "Haumea",
    "Makemake",
    "Bamberga",
    "Patientia",
    "Thisbe",
    "Herculina",
    "Doris",
    "Ursula",
    "Eugenia",
    "Amphitrite",
    "Diotima",
    "Fortuna",
    "Egeria",
    "Themis",
    "Aurora",
    "Alauda",
    "Aletheia",
    "Palma",
    "Nemesis",
    "Psyche",
    "Hebe",
    "Lachesis",
    "Daphne",
    "Bertha",
    "Freia",
    "Winchester",
    "Hilda",
    "Pretoria",
    "Metis",
    "Aegle",
    "Kalliope",
    "Germania",
    "Prokne",
    "Stereoskopia",
    "Agamemnon",
    "Alexandra",
    "Siegena",
    "Elpis",
    "Real Lilith",
    "Black Sun 2",
    "Vulcan",
    "Borasisi",
    "Lempo",
    "1998 SM165",
    "Ceto",
    "Teharonhiawako",
    "2000 OJ67",
    "Elektra",
    "Typhon",
    "Aspasia",
    "Chicago",
    "Loreley",
    "Gyptis",
    "Diomedes",
    "Kreusa",
    "Juewa",
    "Eunike",
    "Ino",
    "Ismene",
    "Merapi"
]

        houses_objects = [
            "house  1",
            "house  2",
            "house  3",
            "house  4",
            "house  5",
            "house  6",
            "Vertex"
        ]
        planets_object = [
            'Sun',
            'Moon',
            'Mercury',
            'Venus',
            'Mars',
            'Jupiter',
            'Saturn',
            'Uranus',
            'Neptune',
            'Pluto',
            'true Node'

        ]
        output = result.stdout
        # First House
        houses_1_parse_output = parse_houses_and_vertex(output,1)
        # # Second House 
        houses_2_parse_output = parse_houses_and_vertex(output,2)
        # # Third House
        houses_3_parse_output = parse_houses_and_vertex(output,3)
        # # Fourth House
        houses_4_parse_output = parse_houses_and_vertex(output,4)
        # # Fifth House
        houses_5_parse_output = parse_houses_and_vertex(output,5)
        # # Sixth House
        houses_6_parse_output = parse_houses_and_vertex(output,6)
        # # Vertex
        houses_vertex_parse_output = parse_houses_and_vertex(output,houses_objects[6])

        # print(houses_parse_output)
        print(f"Data of the Houses: {houses_1_parse_output}")
        


        output2 = result2.stdout
        lines2 = output2.splitlines()



        # # Output of the Planets 
        planet_sun_parse_output = parse_planets(output2,planets_object[0])
        print(f"Data of the Sun: {planet_sun_parse_output}")
        planet_moon_parse_output = parse_planets(output2,planets_object[1])
        print(f"Data of the Moon: {planet_moon_parse_output}")
        planet_mercury_parse_output = parse_planets(output2,planets_object[2])
        print(f"Data of the Mercury: {planet_mercury_parse_output}")
        planet_venus_parse_output = parse_planets(output2,planets_object[3])
        print(f"Data of the Venus: {planet_venus_parse_output}")
        planet_mars_parse_output = parse_planets(output2,planets_object[4])
        print(f"Data of the Mars: {planet_mars_parse_output}")
        planet_jupiter_parse_output = parse_planets(output2,planets_object[5])
        print(f"Data of the Jupiter: {planet_jupiter_parse_output}")
        planet_saturn_parse_output = parse_planets(output2,planets_object[6])
        print(f"Data of the Saturn: {planet_saturn_parse_output}")
        planet_uranus_parse_output = parse_planets(output2,planets_object[7])
        print(f"Data of the Uranus: {planet_uranus_parse_output}")
        planet_neptune_parse_output = parse_planets(output2,planets_object[8])
        print(f"Data of the Neptune: {planet_neptune_parse_output}")
        planet_pluto_parse_output = parse_planets(output2,planets_object[9])
        print(f"Data of the Pluto: {planet_pluto_parse_output}")
        planet_true_node_parse_output = parse_planets(output2,planets_object[10])
        print(f"Data of the True Node: {planet_true_node_parse_output}")
        # Hypothetical Planet 
        lilith_real_planet_result_output = lilith_real_planet_result.stdout
        lilith_real_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'osc. Apogee')
        # Sol Negro 2
        sol_blanco_planet_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'intp. Perigee')

        # Vulcan
        vulcan_planet_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'Vulcan ')

        



        quiron_output = quiron_planet_result.stdout
        quiron_parse_output= parse_asteroid_output(quiron_output,astro_objects[0])
        
        lilith_output = lilith_planet_result.stdout
        lilith_parse_output = parse_asteroid_output(lilith_output,astro_objects[1])

        cerus_output = cerus_planet_result.stdout
        cerus_parse_output = parse_asteroid_output(cerus_output,astro_objects[3])

        pallas_output = pallas_planet_result.stdout
        pallas_parse_output = parse_asteroid_output(pallas_output,astro_objects[4])

        juno_output = juno_planet_result.stdout
        juno_parse_output = parse_asteroid_output(juno_output,astro_objects[5])

        vesta_output = vesta_planet_result.stdout
        vesta_parse_output = parse_asteroid_output(vesta_output,astro_objects[6])

        eris_output = eris_planet_result.stdout
        eris_parse_output = parse_asteroid_output(eris_output,astro_objects[7])

        white_moon_output = white_moon_result.stdout
        white_moon_parse_output = parse_asteroid_output(white_moon_output,astro_objects[8])

        quaoar_output = quaoar_planet_result.stdout
        quaoar_parse_output = parse_asteroid_output(quaoar_output,astro_objects[9])

        sedna_output = sedna_planet_result.stdout
        sedna_parse_output = parse_asteroid_output(sedna_output,astro_objects[10])

        varuna_output = varuna_planet_result.stdout
        varuna_parse_output = parse_asteroid_output(varuna_output,astro_objects[11])

        nessus_output = nessus_planet_result.stdout
        nessus_parse_output = parse_asteroid_output(nessus_output,astro_objects[12])

        waltemath_output = waltemath_planet_result.stdout
        waltemath_parse_output = parse_asteroid_output(waltemath_output,astro_objects[13])

        hygeia_output = hygeia_planet_result.stdout
        hygeia_parse_output = parse_asteroid_output(hygeia_output,astro_objects[14])

        sylvia_output = sylvia_planet_result.stdout
        sylvia_parse_output = parse_asteroid_output(sylvia_output,astro_objects[15])

        hektor_output = hektor_planet_result.stdout
        hektor_parse_output = parse_asteroid_output(hektor_output,astro_objects[16])

        europa_output = europa_planet_result.stdout
        europa_parse_output = parse_asteroid_output(europa_output,astro_objects[17])

        davida_output = davida_planet_result.stdout
        davida_parse_output = parse_asteroid_output(davida_output,astro_objects[18])

        interamnia_output = interamnia_planet_result.stdout
        interamnia_parse_output = parse_asteroid_output(interamnia_output,astro_objects[19])

        camilla_output = camilla_planet_result.stdout
        camilla_parse_output = parse_asteroid_output(camilla_output,astro_objects[20])

        cybele_output = cybele_planet_result.stdout
        cybele_parse_output = parse_asteroid_output(cybele_output,astro_objects[21])

        chariklo_output = chariklo_planet_result.stdout
        chariklo_parse_output = parse_asteroid_output(chariklo_output,astro_objects[29])

        iris_output = iris_planet_result.stdout
        iris_parse_output = parse_asteroid_output(iris_output,astro_objects[30])

        eunomia_planet_output = eunomia_planet_result.stdout
        eunomia_parse_output = parse_asteroid_output(eunomia_planet_output,astro_objects[31])

        euphrosyne_output = euphrosyne_planet_result.stdout
        euphrosyne_parse_output = parse_asteroid_output(euphrosyne_output,astro_objects[32])

        orcus_output = orcus_planet_result.stdout
        orcus_parse_output = parse_asteroid_output(orcus_output,astro_objects[33])

        pholus_output = pholus_planet_result.stdout
        pholus_parse_output = parse_asteroid_output(pholus_output,astro_objects[34])

        hermione_output = hermione_planet_result.stdout
        hermione_parse_output = parse_asteroid_output(hermione_output,astro_objects[35])

        ixion_output = ixion_planet_result.stdout
        ixion_parse_output = parse_asteroid_output(ixion_output,astro_objects[36])

        haumea_output = haumea_planet_result.stdout
        haumea_parse_output = parse_asteroid_output(haumea_output,astro_objects[37])

        makemake_output = makemake_planet_result.stdout
        makemake_parse_output = parse_asteroid_output(makemake_output,astro_objects[38])

        bamberga_output = bamberga_planet_result.stdout
        bamberga_parse_output = parse_asteroid_output(bamberga_output,astro_objects[39])

        patientia_output = patientia_planet_result.stdout
        patientia_parse_output = parse_asteroid_output(patientia_output,astro_objects[40])

        thisbe_output = thisbe_planet_result.stdout
        thisbe_parse_output = parse_asteroid_output(thisbe_output,astro_objects[41])

        herculina_output = herculina_planet_result.stdout
        herculina_parse_output = parse_asteroid_output(herculina_output,astro_objects[42])

        doris_output = doris_planet_result.stdout
        doris_parse_output = parse_asteroid_output(doris_output,astro_objects[43])  

        ursula_output = ursula_planet_result.stdout
        ursula_parse_output = parse_asteroid_output(ursula_output,astro_objects[44])

        eugenia_output = eugenia_planet_result.stdout
        eugenia_parse_output = parse_asteroid_output(eugenia_output,astro_objects[45])

        amphitrite_output = amphitrite_planet_result.stdout
        amphitrite_parse_output = parse_asteroid_output(amphitrite_output,astro_objects[46])

        diotima_output = diotima_planet_result.stdout
        diotima_parse_output = parse_asteroid_output(diotima_output,astro_objects[47])

        fortuna_output = fortuna_planet_result.stdout
        fortuna_parse_output = parse_asteroid_output(fortuna_output,astro_objects[48])

        egeria_output = egeria_planet_result.stdout
        egeria_parse_output = parse_asteroid_output(egeria_output,astro_objects[49])

        themis_output = themis_planet_result.stdout
        themis_parse_output = parse_asteroid_output(themis_output,astro_objects[50])

        aurora_output = aurora_planet_result.stdout
        aurora_parse_output = parse_asteroid_output(aurora_output,astro_objects[51])

        alauda_output = alauda_planet_result.stdout
        alauda_parse_output = parse_asteroid_output(alauda_output,astro_objects[52])

        aletheia_output = aletheia_planet_result.stdout
        aletheia_parse_output = parse_asteroid_output(aletheia_output,astro_objects[53])

        palma_output = palma_planet_result.stdout
        palma_parse_output = parse_asteroid_output(palma_output,astro_objects[54])

        nemesis_output = nemesis_planet_result.stdout
        nemesis_parse_output = parse_asteroid_output(nemesis_output,astro_objects[55])

        psyche_output = psyche_planet_result.stdout
        psyche_parse_output = parse_asteroid_output(psyche_output,astro_objects[56])

        hebe_output = hebe_planet_result.stdout
        hebe_parse_output = parse_asteroid_output(hebe_output,astro_objects[57])

        lachesis_output = lachesis_planet_result.stdout
        lachesis_parse_output = parse_asteroid_output(lachesis_output,astro_objects[58])

        daphne_output = daphne_planet_result.stdout
        daphne_parse_output = parse_asteroid_output(daphne_output,astro_objects[59])

        bertha_output = bertha_planet_result.stdout
        bertha_parse_output = parse_asteroid_output(bertha_output,astro_objects[60])

        freia_output = freia_planet_result.stdout
        freia_parse_output = parse_asteroid_output(freia_output,astro_objects[61])

        winchester_output = winchester_planet_result.stdout
        winchester_parse_output = parse_asteroid_output(winchester_output,astro_objects[62])
        

        hilda_output = hilda_planet_result.stdout
        hilda_parse_output = parse_asteroid_output(hilda_output,astro_objects[63])

        pretoria_output = pretoria_planet_result.stdout
        pretoria_parse_output = parse_asteroid_output(pretoria_output,astro_objects[64])

        metis_output = metis_planet_result.stdout
        metis_parse_output = parse_asteroid_output(metis_output,astro_objects[65])

        aegle_output = aegle_planet_result.stdout
        aegle_parse_output = parse_asteroid_output(aegle_output,astro_objects[66])

        kalliope_output = kalliope_planet_result.stdout
        kalliope_parse_output = parse_asteroid_output(kalliope_output,astro_objects[67])

        germania_output = germania_planet_result.stdout
        germania_parse_output = parse_asteroid_output(germania_output,astro_objects[68])

        prokne_output = prokne_planet_result.stdout
        prokne_parse_output = parse_asteroid_output(prokne_output,astro_objects[69])

        stereoskopia_output = stereoskopia_planet_result.stdout
        stereoskopia_parse_output = parse_asteroid_output(stereoskopia_output,astro_objects[70])

        agamemnon_output = agamemnon_planet_result.stdout
        agamemnon_parse_output = parse_asteroid_output(agamemnon_output,astro_objects[71])

        alexandra_output = alexandra_planet_result.stdout
        alexandra_parse_output = parse_asteroid_output(alexandra_output,astro_objects[72])

        siegena_output = siegena_planet_result.stdout
        siegena_parse_output = parse_asteroid_output(siegena_output,astro_objects[73])

        elpis_output = elpis_planet_result.stdout
        elpis_parse_output = parse_asteroid_output(elpis_output,astro_objects[74])

        borasisi_output = borasisi_planet_result.stdout
        borasisi_parse_output = parse_asteroid_output(borasisi_output,astro_objects[78])

        lempo_output = lempo_planet_result.stdout
        lempo_parse_output = parse_asteroid_output(lempo_output,astro_objects[79])

        _1998_26308_output = _1998_26308_planet_result.stdout
        _1998_26308_parse_output = parse_asteroid_output(_1998_26308_output,astro_objects[80])

        ceto_output = ceto_planet_result.stdout
        ceto_parse_output = parse_asteroid_output(ceto_output,astro_objects[81])

        teharonhiawako_output = teharonhiawako_planet_result.stdout
        teharonhiawako_parse_output = parse_asteroid_output(teharonhiawako_output,astro_objects[82])

        _2000_oj67_134860_output = _2000_oj67_134860_planet_result.stdout
        _2000_oj67_134860_parse_output = parse_asteroid_output(_2000_oj67_134860_output,astro_objects[83])

        elektra_output = elektra_planet_result.stdout
        elektra_parse_output = parse_asteroid_output(elektra_output,astro_objects[84])

        typhon_output = typhon_planet_result.stdout
        typhon_parse_output = parse_asteroid_output(typhon_output,astro_objects[85])

        aspasia_output = aspasia_planet_result.stdout
        aspasia_parse_output = parse_asteroid_output(aspasia_output,astro_objects[86])

        chicago_output = chicago_planet_result.stdout
        chicago_parse_output = parse_asteroid_output(chicago_output,astro_objects[87])

        loreley_output = loreley_planet_result.stdout
        loreley_parse_output = parse_asteroid_output(loreley_output,astro_objects[88])

        gyptis_output = gyptis_planet_result.stdout
        gyptis_parse_output = parse_asteroid_output(gyptis_output,astro_objects[89])

        diomedes_output = diomedes_planet_result.stdout
        diomedes_parse_output = parse_asteroid_output(diomedes_output,astro_objects[90])

 

        kreusa_output = kreusa_planet_result.stdout
        kreusa_parse_output = parse_asteroid_output(kreusa_output,astro_objects[91])

        juewa_output = juewa_planet_result.stdout
        juewa_parse_output = parse_asteroid_output(juewa_output,astro_objects[92])

        eunike_output = eunike_planet_result.stdout
        eunike_parse_output = parse_asteroid_output(eunike_output,astro_objects[93])

        ino_output = ino_planet_result.stdout
        ino_parse_output = parse_asteroid_output(ino_output,astro_objects[94])

        ismene_output = ismene_planet_result.stdout
        ismene_parse_output = parse_asteroid_output(ismene_output,astro_objects[95])

        merapi_output = merapi_planet_result.stdout
        merapi_parse_output = parse_asteroid_output(merapi_output,astro_objects[96])



        # Create a dictionary to store the result data that are Empty in the Excel
        sol_negro_parse_output =  {
            "name": "Sol Negro",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
            "retrograde": ""
        }
        # For AntiVertex
        anti_vertex_parse_output = {
             "name": "Antivertex",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "" ,
             "retrograde": ""
        }
        # For Nodo Sur Real
        nodo_sur_real_parse_output = {
               "name": "Nodo Sur Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""  
        }
        # For Sol Negro Real
        sol_negro_real_parse_output = {
            "name": "Sol Negro Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # For Lilith 2
        lilith2_parse_output = {
            "name": "Lilith 2",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # For Waldemath Priapus
        waltemath_priapus_parse_output = {
            "name": "Waldemath Priapus",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # Sol Blanco 
        sol_blanco_parse_output = {
            "name": "Sol Blanco",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        } 

        planets = []



        # Open the workbook outside of the loop to avoid repeated opening and closing
        try:
            original_path = r'C:\El Camino que Creas\Generador de Informes\Generador de Informes\Generador de Informes.xlsm'
            # base, ext = os.path.splitext(original_path)
            # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f_natal_chart")  # Format: YYYYMMDD_HHMMSS_milliseconds
            # copied_file_path = f"{base}_{timestamp}{ext}"
            # # wb = xl.Workbooks.Open(file_path)  # Path to your Excel file
            # shutil.copyfile(original_path, copied_file_path)
            wb = xl.Workbooks.Open(original_path)  # Path to your Excel file
            
            try:
                
                sheet_name = 'CN y RS (o RL)'  # Replace with your sheet name
                sheet = wb.Sheets(sheet_name)

                # List of Data for Natal Positions
                asteroidsList = [houses_1_parse_output,houses_2_parse_output,houses_3_parse_output,houses_4_parse_output,houses_5_parse_output,houses_6_parse_output,planet_sun_parse_output,planet_moon_parse_output,planet_mercury_parse_output,planet_venus_parse_output,planet_mars_parse_output,planet_jupiter_parse_output,planet_saturn_parse_output,planet_uranus_parse_output,planet_neptune_parse_output,planet_pluto_parse_output,planet_true_node_parse_output,quiron_parse_output,lilith_parse_output,houses_vertex_parse_output,cerus_parse_output,pallas_parse_output,juno_parse_output,vesta_parse_output,eris_parse_output,white_moon_parse_output,quaoar_parse_output,sedna_parse_output,varuna_parse_output,nessus_parse_output,waltemath_parse_output,hygeia_parse_output,sylvia_parse_output,hektor_parse_output,europa_parse_output,davida_parse_output,interamnia_parse_output,camilla_parse_output,cybele_parse_output,sol_negro_parse_output,anti_vertex_parse_output,nodo_sur_real_parse_output,sol_negro_real_parse_output,lilith2_parse_output,waltemath_priapus_parse_output,sol_blanco_parse_output,chariklo_parse_output,iris_parse_output,eunomia_parse_output,euphrosyne_parse_output,orcus_parse_output,pholus_parse_output,hermione_parse_output,ixion_parse_output,haumea_parse_output,makemake_parse_output,bamberga_parse_output,patientia_parse_output,thisbe_parse_output,herculina_parse_output,doris_parse_output,ursula_parse_output,eugenia_parse_output,amphitrite_parse_output,diotima_parse_output,fortuna_parse_output,egeria_parse_output,themis_parse_output,aurora_parse_output,alauda_parse_output,aletheia_parse_output,palma_parse_output,nemesis_parse_output,psyche_parse_output,hebe_parse_output,lachesis_parse_output,daphne_parse_output,bertha_parse_output,freia_parse_output,winchester_parse_output,hilda_parse_output,pretoria_parse_output,metis_parse_output,aegle_parse_output,kalliope_parse_output,germania_parse_output,prokne_parse_output,stereoskopia_parse_output,agamemnon_parse_output,alexandra_parse_output,siegena_parse_output,elpis_parse_output,lilith_real_parse_output,sol_blanco_planet_parse_output,vulcan_planet_parse_output,borasisi_parse_output,lempo_parse_output,_1998_26308_parse_output,ceto_parse_output,teharonhiawako_parse_output,_2000_oj67_134860_parse_output,elektra_parse_output,typhon_parse_output,aspasia_parse_output,chicago_parse_output,loreley_parse_output,gyptis_parse_output,diomedes_parse_output,kreusa_parse_output,juewa_parse_output,eunike_parse_output,ino_parse_output,ismene_parse_output,merapi_parse_output]
                
                
                
                date_sun_report = datetime.strptime(sun_return_date, "%Y-%m-%d")
                solar_return_date = get_solar_return_data(birth_date_year,birth_date_month,birth_date_day,ut_hour,ut_min,ut_sec,date_sun_report.year,date_sun_report.month,date_sun_report.day)
                print("Getting Solar Return Date Shahryar %s" % solar_return_date.get('solar_return_date_start'))
                # Solar Return Position Date
                get_solar_return_position = get_solar_return_position_func(lat_deg,lon_deg,report_type_data,solar_return_date.get('solar_return_date_start')) 
                print("Getting Solar Return Position Data Shahryar %s" % get_solar_return_position)
                start_row = 5  # Row 29
                start_column = 3  # Column S
                for index, asteroid in enumerate(asteroidsList):
                 row = start_row + index
                #  sheet.Cells(row, start_column).Value = asteroid['name']
                 sheet.Cells(row, start_column + 1).Value = asteroid['position_sign']
               
                 sheet.Cells(row, start_column + 2).Value =asteroid['positionDegree']
                 sheet.Cells(row, start_column + 3).Value =  asteroid['position_min']
                 sheet.Cells(row, start_column + 4).Value = asteroid['position_sec']
                 sheet.Cells(row, start_column + 5).Value = asteroid['retrograde']
                #  In 19 Col Row 5
                # Put the name of the User
                sheet.Cells(5,19).Value = person_name
                # In Row 8 Local Birth date with seconds.
                sheet.Cells(8,19).Value = person_birth_date_local
                 
                # In Row 11 Put the User Location. 
                sheet.Cells(11,19).Value = person_location
                # In Row 14 Put Natal Chart | Sun Return | Moon Return
                sheet.Cells(14,19).Value = report_type_data 
                # Sun Return Date
                    # Sun Return Date In Row 17
                sheet.Cells(17,19).Value = solar_return_date.get('solar_return_date_start')
                    # Moon Return Date In Row 20
                sheet.Cells(20,19).Value = solar_return_date.get('solar_return_date_end')
                    # Gender In 21 kah 5 


                sun_col = 11
                for index, sun_return_asteroids in enumerate(get_solar_return_position):
                 row = start_row + index
                #  sheet.Cells(row, sun_col).Value = sun_return_asteroids['name']
                 sheet.Cells(row, sun_col + 1).Value = sun_return_asteroids['position_sign']
               
                 sheet.Cells(row, sun_col + 2).Value =sun_return_asteroids['positionDegree']
                 sheet.Cells(row, sun_col + 3).Value =  sun_return_asteroids['position_min']
                 sheet.Cells(row, sun_col + 4).Value = sun_return_asteroids['position_sec']
                 sheet.Cells(row, sun_col + 5).Value = sun_return_asteroids['retrograde']

                sheet.Cells(5,21).Value = gender_type




                print("Data modified successfully.")
                return jsonify({"message": "Data modified successfully.", "result2": planets, "asteriods": asteroidsList,"fileName":original_path,"other_list":get_solar_return_position}), 200
            finally:
                wb.Close(SaveChanges=True)  # Save changes after running macro
        except Exception as e:
            print("Error opening workbook:", e)
            return jsonify({"error": str(e)}), 500
        finally:
            xl.Quit()
   
   
    except Exception as e:
        print("Error initializing Excel:", e)
        logger.error(f"Error occurred: {str(e)}\n{traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

def parse_asteroid_output(asteroid_pholus_output,asteroid_object_name):
    lines = asteroid_pholus_output.splitlines()  # Split by newline characters
    result = {}
    
    
    try:
        if len(lines) > 0:
            pattern2 = re.escape(asteroid_object_name)
            parts = re.split(pattern2, asteroid_pholus_output)
            pattern = r'[a-zA-Z]'
            match = re.split(pattern, parts[1])
            # Remove Extra spaces. 
            removeExtraSpaceDegree = re.sub(r'\s+', '', match[0])
            first_two_alphabets = []

            

# Iterate through the characters in the string
            for char in parts[1]:
                if char.isalpha():  # Check if the character is alphabetic
                    first_two_alphabets.append(char)
                    if len(first_two_alphabets) == 2:  # Stop once we have two alphabetic characters
                        break

            # degree_match_sign = re.findall(r'[a-zA-Z]+', match)

            first_two_alphabets = ''.join(first_two_alphabets)
            splitbySign = re.split(first_two_alphabets, parts[1])
            # Remove "\n"
            removeNewLine = re.sub(r'\n', '', splitbySign[1])
            # Remove Extra Qutotation
            removeExtraQuotation = re.sub(r'"', '', removeNewLine)
            pattern3 = r'\s{3,}'  # Pattern to split by 2 or more spaces
            splitbyTwoSpaces = re.split(pattern3, removeExtraQuotation)
            removedExtraSpaces = splitbyTwoSpaces[0].replace(" ", "")
            # Split by ' 
            splitbySingleQuote = re.split("'", removedExtraSpaces)
            print(parts)
            # Split by °
            splitbyDegree = re.split("Â°", splitbyTwoSpaces[1])
            # Assuming splitbyDegree[0] is a string and you need to convert it to an integer
            value = int(splitbyDegree[0])

            # Check the condition and return "R" or an empty string
            resultValue = "R" if '-' in splitbyDegree[0] else ""

            

            result[asteroid_object_name] = {
                    "name" : asteroid_object_name,
                    "positionDegree": removeExtraSpaceDegree,
                    "position_sign": zodiac_signs.get(first_two_alphabets.lower(), first_two_alphabets),
                    "position_min":splitbySingleQuote[0],
                    "position_sec":splitbySingleQuote[1],
                    "retrograde": resultValue
                
                    # "commands": lines,              
                
    
            }
        else:
            result["error"] = "Error parsing output: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result[asteroid_object_name]  # Always return a dictionary

def parse_houses_and_vertex(asteroid_pholus_output, house_number):
    # Split by newline characters
    lines = asteroid_pholus_output.splitlines()  
    print(f"Lines: {lines}")
    
    # Initialize result dictionary
    result = {}

    # Loop through each line to find the house of interest
    for line in lines:
        if f"house  {house_number}" in line:
            # Remove "house <house_number>" from the line
            output_string = re.sub(rf"house\s+{house_number}\s+", "", line)
            print(f"Output String after regex: {output_string}")
            # Remove the Spaces
            # output_space_removed = output_string.replace
            # Split by 2 
            pattern = r'[a-zA-Z]'
            # Split by 2 alphabets
            splitByAlphaBets = re.split(pattern, output_string)
            print(f"AlphaBets String after regex: {splitByAlphaBets}")    
            # Split by spaces
            splitbySpace = re.split(r'\s+', output_string)
            # positionDegree = splitbySpace[0]
            positionSign = splitbySpace[1]
            # positionMinSec = splitbySpace[2]
            
            # Separate by single quote and remove double quotes
            splitbySingleQuote = re.split("'", splitByAlphaBets[2])
            splitbySingleQuote[1] = splitbySingleQuote[1].replace('"', "")
            
            result = {
                "name": f"Casa {house_number}",
                "positionDegree": splitByAlphaBets[0].replace(" ", ""),
                "position_sign": zodiac_signs.get(positionSign.lower(), positionSign),
                "position_min": splitbySingleQuote[0].replace(" ", ""),
                "position_sec": splitbySingleQuote[1].replace(" ", ""),
                "retrograde": "",
            }
            break  # Exit the loop after finding the required house
        elif "Vertex" in line:
            output_string = re.sub(rf"Vertex\s+", "", line)
            pattern = r'[a-zA-Z]'
            # Split by 2 alphabets
            splitByAlphaBets = re.split(pattern, output_string)
            print(f"AlphaBets String after regex: {splitByAlphaBets}")    
            # Split by spaces
            splitbySpace = re.split(r'\s+', output_string)
            # positionDegree = splitbySpace[0]
            positionSign = splitbySpace[1]
            # positionMinSec = splitbySpace[2]
            
            # Separate by single quote and remove double quotes
            splitbySingleQuote = re.split("'", splitByAlphaBets[2])
            splitbySingleQuote[1] = splitbySingleQuote[1].replace('"', "")

            print(f"Output String after regex: {line}")
            result ={
                "name": f"Vertex",
                "positionDegree": splitByAlphaBets[0].replace(" ", ""),
                "position_sign": zodiac_signs.get(positionSign.lower(), positionSign),
                "position_min": splitbySingleQuote[0].replace(" ", ""),
                "position_sec": splitbySingleQuote[1].replace(" ", ""),
                "retrograde": "",
                # "commands": lines
            }
    return result

def parse_planets(planets_output, planet_name):
    # Split by newline characters
    lines = planets_output.splitlines()
    # Loop through each line to find the planet of interest
    # Initialize result dictionary
    result = {}
    for line in lines:
        if planet_name in line:
            # Remove the planet name from the line
            output_string = re.sub(rf"{planet_name}\s+", "", line)
            # Split by 2 spaces
            splitbySpace = re.split(r'\s{3,}', output_string)


            # Split by Â°
            splitbyDegree = re.split("Â°", splitbySpace[1])
            # speed
            speed = splitbyDegree[0]
            outputData = splitbySpace[0]
             # Remove the Spaces 
            outputData = outputData.replace(" ", "")
            # Split by alphabets
            pattern = r'[a-zA-Z]'
            splitByAlphaBets = re.split(pattern, outputData)
            # Degree
            degree = splitByAlphaBets[0]
            # minsec
            minsec = splitByAlphaBets[2]
            # Split by single quote
            splitbySingleQuote = re.split("'", minsec)
        

            
            # Find the 2 alphabets in outputData
            first_two_alphabets = []
            # Iterate through the characters in the string
            for char in outputData:
                if char.isalpha():  # Check if the character is alphabetic
                    first_two_alphabets.append(char)
                    if len(first_two_alphabets) == 2:  # Stop once we have two alphabetic characters
                        break

            resultValue = resultValue = "R" if '-' in speed else ""
            positionSign = first_two_alphabets[0]+first_two_alphabets[1]
             


            result = {
                        "name": planet_name,
                        "positionDegree": degree,
                        "position_sign": zodiac_signs.get(positionSign.lower(), positionSign),
                        "position_min": splitbySingleQuote[0],
                        "position_sec": splitbySingleQuote[1],
                        "retrograde": resultValue      
                        }
            
            



           
        
    # print(f"Lines: {lines}")


    return result

def get_solar_return_data(birth_date_year, birth_date_month, birth_date_day, ut_hour, ut_min, ut_sec, user_selected_year, user_selected_month, user_selected_day):
    # Solar Return Date 
    find_solar_return_date = datetime(user_selected_year, user_selected_month, user_selected_day, ut_hour, ut_min, ut_sec)
    
    # Get the Julian Day for the birth date and time
    jd_birth = swe.julday(birth_date_year, birth_date_month, birth_date_day, ut_hour + ut_min / 60 + ut_sec / 3600)
    
    # Get the Sun position at birth
    sun_pos, ret = swe.calc_ut(jd_birth, swe.SUN)
    birth_sun_longitude = sun_pos[0]
    
    # Estimate Julian Day for the solar return (close to the birthday)
    jd_estimate = swe.julday(user_selected_year, user_selected_month, user_selected_day)
    
    # Time delta for Solar Return (similar to time_delta_moon but adapted for solar return context)
    time_delta_sun = timedelta(days=365)  # Typical interval for solar returns is approximately a year
    
    # Find the exact time the Sun returns to the same longitude using solcross_ut
    serr = ''
    jd_solar_return = swe.solcross_ut(birth_sun_longitude, jd_estimate, 0)
    
    if jd_solar_return < jd_estimate:
        return {'error': serr}, 400
    
    # Convert Julian Day to calendar date and time
    solar_return_date = swe.revjul(jd_solar_return)
    solar_return_date_datetime = datetime(solar_return_date[0], solar_return_date[1], solar_return_date[2],
                                          int(solar_return_date[3]), int((solar_return_date[3] % 1) * 60),
                                          int(((solar_return_date[3] % 1) * 60 % 1) * 60))
    
    print(f"Getting Solar Return Date {solar_return_date_datetime}")
    
    # Now find the difference between solar_return_date_datetime and find_solar_return_date
    difference = find_solar_return_date - solar_return_date_datetime
    
    # Check how many times of time_delta_sun is in difference
    times_difference_times = difference / time_delta_sun
    print(f"Times Difference {times_difference_times}")
    print(f"Difference {difference}")
    
    # Adjust solar return date if needed
    if difference > time_delta_sun:
        # Find the next solar return
        jd_solar_return = swe.solcross_ut(birth_sun_longitude, jd_solar_return + 365 * (times_difference_times - 1), 0)
        
        # Convert Julian Day to calendar date and time
        solar_return_date = swe.revjul(jd_solar_return)
        solar_return_date_datetime = datetime(solar_return_date[0], solar_return_date[1], solar_return_date[2],
                                              int(solar_return_date[3]), int((solar_return_date[3] % 1) * 60),
                                              int(((solar_return_date[3] % 1) * 60 % 1) * 60))
        print(f"Getting Next Solar Return Date {solar_return_date_datetime}")
    if difference < time_delta_sun:
        # difference value 
        # Remove the sign of the difference
        const_value = 365 * (abs(times_difference_times) + 1)
        print(f"Const Value {const_value}")
        jd_solar_return = swe.solcross_ut(birth_sun_longitude, jd_solar_return - const_value, 0)
        
        # Convert Julian Day to calendar date and time
        solar_return_date = swe.revjul(jd_solar_return)
        solar_return_date_datetime = datetime(solar_return_date[0], solar_return_date[1], solar_return_date[2],
                                              int(solar_return_date[3]), int((solar_return_date[3] % 1) * 60),
                                              int(((solar_return_date[3] % 1) * 60 % 1) * 60))
        print(f"Start Date Moon Cross User {solar_return_date_datetime}")
        # Add the timedelta to the datetime object
        end_date_moon_cross_user = solar_return_date_datetime + time_delta_sun
        print(f"End Date Moon Cross User {end_date_moon_cross_user}")
    
    # Add the timedelta to the datetime object
    end_date_sun_cross_user = solar_return_date_datetime + time_delta_sun
    print(f"Start Date Sun Cross User {solar_return_date_datetime}")
    print(f"End Date Sun Cross User {end_date_sun_cross_user}")
    
    return {
        "solar_return_date_start": solar_return_date_datetime.strftime("%Y/%m/%d %H:%M:%S"),
        "solar_return_date_end": end_date_sun_cross_user.strftime("%Y/%m/%d %H:%M:%S")
    }

    # Get the Solar Return Chart With Respect of get_solar_return_data
def get_solar_return_position_func(lat_deg,lon_deg,report_type_data,date):
    print("Getting Solar Return Position %s" % date)
        
    # Split the date and time
    date_part, time_part = date.split(" ")
    
    # Further split the date and time into components
    year, month, day = date_part.split("/")
    hour, minute, second = time_part.split(":")
    
    # Print each component
    print("Year:", year)
    print("Month:", month)
    print("Day:", day)
    print("Hour:", hour)
    print("Minute:", minute)
    print("Second:", second)
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Get the parameters from the request data and ensure they are integers
        birth_date_year = int(year)
        birth_date_month = int(month)
        birth_date_day = int(day)
        ut_hour = int(hour)
        ut_min = int(minute)
        ut_sec = int(second)
        lat_deg = lat_deg
        lon_deg = lon_deg
        # Moon Return, Solar Return or Natal return 
        report_type_data = report_type_data

        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False  # Set to True if you want Excel to be visible

        # Construct the command with zero-padded values
        # For House Data From Cell D5 to D10
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
        # For Planets Data From Cell D11 to D21 Which Includes True Node
        command2 = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZS -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"
        # For Quirón Command From Cell D22sky
        quiron_planet = f"swetest -ps -xs2060 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Lilith Command From Cell D23
        lilith_planet = f"swetest -ps -xs1181 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Cerus Command 
        cerus_planet = f"swetest -ps -xs1 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Pallas Command
        pallas_planet = f"swetest -ps -xs2 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Juno Command
        juno_planet = f"swetest -ps -xs3 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # For Vesta Command
        vesta_planet = f"swetest -ps -xs4 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Eris Command
        eris_planet = f"swetest -ps -xs136199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # White Moon Command
        white_moon = f"swetest -pZ -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Quaoar Command
        quaoar_planet = f"swetest -ps -xs50000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Sedna Command
        sedna_planet = f"swetest -ps -xs90377 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Varuna Command
        varuna_planet = f"swetest -ps -xs20000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Nessus Command
        nessus_planet = f"swetest -ps -xs7066 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Waltemath Command
        waltemath_planet = f"swetest -pw -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hygeia Command
        hygeia_planet = f"swetest -ps -xs10 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Sylvia Command
        sylvia_planet = f"swetest -ps -xs87 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
#         Hektor	624	Hector
        hektor_planet = f"swetest -ps -xs624 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Europa	52	Europa
        europa_planet = f"swetest -ps -xs52 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Davida	511	Davida
        davida_planet = f"swetest -ps -xs511 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Interamnia	704	Interamnia
        interamnia_planet = f"swetest -ps -xs704 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Camilla	107	Camilla
        camilla_planet = f"swetest -ps -xs107 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Cybele	65	Cybele
        cybele_planet = f"swetest -ps -xs65 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Sol Negro	h22	Black Sun

# Antivertex		Anti-Vertex

# Nodo Sur Real		True South Node
# Sol Negro Real		True Black Sun
# Lilith 2		Lilith 2
# Waldemath Priapus		Waldemath Priapus
# Sol Blanco		White Sun
# Chariklo	10199	Chariklo
        chariklo_planet = f"swetest -ps -xs10199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Iris	7	Iris
        iris_planet = f"swetest -ps -xs7 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Eunomia	15	Eunomia
        eunomia_planet = f"swetest -ps -xs15 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Euphrosyne	31	Euphrosyne
        euphrosyne_planet = f"swetest -ps -xs31 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Orcus	90482	Orcus
        # Orcus Command
        orcus_planet = f"swetest -ps -xs90482 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Pholus	5145	Pholus
        # Pholus Command
        pholus_planet = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hermíone Command
        hermione_planet = f"swetest -ps -xs121 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ixion Command
        ixion_planet = f"swetest -ps -xs28978 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Haumea Command
        haumea_planet = f"swetest -ps -xs136108 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Makemake Command
        makemake_planet = f"swetest -ps -xs136472 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Bamberga Command
        bamberga_planet = f"swetest -ps -xs324 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Patientia Command
        patientia_planet = f"swetest -ps -xs451 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Thisbe Command
        thisbe_planet = f"swetest -ps -xs88 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Herculina Command
        herculina_planet = f"swetest -ps -xs532 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Doris Command
        doris_planet = f"swetest -ps -xs48 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ursula Command
        ursula_planet = f"swetest -ps -xs375 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Eugenia Command
        eugenia_planet = f"swetest -ps -xs45 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Amphitrite Command
        amphitrite_planet = f"swetest -ps -xs29 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Diotima Command
        diotima_planet = f"swetest -ps -xs423 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Fortuna Command
        fortuna_planet = f"swetest -ps -xs19 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Egeria Command
        egeria_planet = f"swetest -ps -xs13 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Themis Command
        themis_planet = f"swetest -ps -xs24 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aurora Command
        aurora_planet = f"swetest -ps -xs94 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Alauda Command
        alauda_planet = f"swetest -ps -xs702 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aletheia Command
        aletheia_planet = f"swetest -ps -xs259 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Palma Command
        palma_planet = f"swetest -ps -xs372 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Nemesis Command
        nemesis_planet = f"swetest -ps -xs128 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Psyche Command
        psyche_planet = f"swetest -ps -xs16 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hebe Command
        hebe_planet = f"swetest -ps -xs6 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lachesis Command
        lachesis_planet = f"swetest -ps -xs120 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Daphne Command
        daphne_planet = f"swetest -ps -xs41 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Bertha Command
        bertha_planet = f"swetest -ps -xs154 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Freia Command
        freia_planet = f"swetest -ps -xs76 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Winchester Command
        winchester_planet = f"swetest -ps -xs747 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Hilda Command
        hilda_planet = f"swetest -ps -xs153 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Pretoria Command
        pretoria_planet = f"swetest -ps -xs790 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Metis Command
        metis_planet = f"swetest -ps -xs9 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aegle Command
        aegle_planet = f"swetest -ps -xs96 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Kalliope Command
        kalliope_planet = f"swetest -ps -xs22 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Germania Command
        germania_planet = f"swetest -ps -xs241 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Prokne Command
        prokne_planet = f"swetest -ps -xs194 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Stereoskopia Command
        stereoskopia_planet = f"swetest -ps -xs566 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Agamemnon Command
        agamemnon_planet = f"swetest -ps -xs911 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Alexandra Command
        alexandra_planet = f"swetest -ps -xs54 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Siegena Command
        siegena_planet = f"swetest -ps -xs386 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Elpis Command
        elpis_planet = f"swetest -ps -xs59 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lilith Real Command
        # osc. Apogee Command in pa
        lilith_real_planet = f"swetest -pa -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Borasisi Command
        borasisi_planet = f"swetest -ps -xs66652 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Lempo Command
        lempo_planet = f"swetest -ps -xs47171 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # 1998(26308) Command
        _1998_26308_planet = f"swetest -ps -xs26308 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Ceto Command
        ceto_planet = f"swetest -ps -xs65489 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Teharonhiawako Command
        teharonhiawako_planet = f"swetest -ps -xs88611 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # 2000 OJ67 (134860) Command
        _2000_oj67_134860_planet = f"swetest -ps -xs134860 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Elektra Command
        elektra_planet = f"swetest -ps -xs130 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Typhon Command
        typhon_planet = f"swetest -ps -xs42355 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Aspasia Command
        aspasia_planet = f"swetest -ps -xs409 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Chicago Command
        chicago_planet = f"swetest -ps -xs334 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Loreley Command
        loreley_planet = f"swetest -ps -xs165 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Gyptis Command
        gyptis_planet = f"swetest -ps -xs444 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Diomedes Command
        diomedes_planet = f"swetest -ps -xs1437 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
        # Kreusa Command
        kreusa_planet = f"swetest -ps -xs488 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Juewa	139	Juewa
        juewa_planet = f"swetest -ps -xs139 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Eunike	185	Eunike
        eunike_planet = f"swetest -ps -xs185 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Ino	173	Ino
        ino_planet = f"swetest -ps -xs173 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Ismene	190	Ismene
        ismene_planet = f"swetest -ps -xs190 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"
# Merapi	536	Merapi
        merapi_planet = f"swetest -ps -xs536 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZS -roundsec"


        # Execute the command using subprocess
        # Planet Names
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        # Houses Names
        result2 = subprocess.run(command2, shell=True, check=True, capture_output=True, text=True)
        quiron_planet_result = subprocess.run(quiron_planet, shell=True, check=True, capture_output=True, text=True)
        lilith_planet_result = subprocess.run(lilith_planet, shell=True, check=True, capture_output=True, text=True)
        cerus_planet_result = subprocess.run(cerus_planet, shell=True, check=True, capture_output=True, text=True)
        pallas_planet_result = subprocess.run(pallas_planet, shell=True, check=True, capture_output=True, text=True)
        juno_planet_result = subprocess.run(juno_planet, shell=True, check=True, capture_output=True, text=True)
        vesta_planet_result = subprocess.run(vesta_planet, shell=True, check=True, capture_output=True, text=True)
        eris_planet_result = subprocess.run(eris_planet, shell=True, check=True, capture_output=True, text=True)
        white_moon_result = subprocess.run(white_moon, shell=True, check=True, capture_output=True, text=True)
        quaoar_planet_result = subprocess.run(quaoar_planet, shell=True, check=True, capture_output=True, text=True)
        sedna_planet_result = subprocess.run(sedna_planet, shell=True, check=True, capture_output=True, text=True)
        varuna_planet_result = subprocess.run(varuna_planet, shell=True, check=True, capture_output=True, text=True)
        nessus_planet_result = subprocess.run(nessus_planet, shell=True, check=True, capture_output=True, text=True)
        waltemath_planet_result = subprocess.run(waltemath_planet, shell=True, check=True, capture_output=True, text=True)
        hygeia_planet_result = subprocess.run(hygeia_planet, shell=True, check=True, capture_output=True, text=True)
        sylvia_planet_result = subprocess.run(sylvia_planet, shell=True, check=True, capture_output=True, text=True)
        hektor_planet_result = subprocess.run(hektor_planet, shell=True, check=True, capture_output=True, text=True)
        europa_planet_result = subprocess.run(europa_planet, shell=True, check=True, capture_output=True, text=True)
        davida_planet_result = subprocess.run(davida_planet, shell=True, check=True, capture_output=True, text=True)
        interamnia_planet_result = subprocess.run(interamnia_planet, shell=True, check=True, capture_output=True, text=True)
        camilla_planet_result = subprocess.run(camilla_planet, shell=True, check=True, capture_output=True, text=True)
        cybele_planet_result = subprocess.run(cybele_planet, shell=True, check=True, capture_output=True, text=True)
        chariklo_planet_result = subprocess.run(chariklo_planet, shell=True, check=True, capture_output=True, text=True)
        iris_planet_result = subprocess.run(iris_planet, shell=True, check=True, capture_output=True, text=True)
        eunomia_planet_result = subprocess.run(eunomia_planet, shell=True, check=True, capture_output=True, text=True)
        euphrosyne_planet_result = subprocess.run(euphrosyne_planet, shell=True, check=True, capture_output=True, text=True)
        orcus_planet_result = subprocess.run(orcus_planet, shell=True, check=True, capture_output=True, text=True)
        pholus_planet_result = subprocess.run(pholus_planet, shell=True, check=True, capture_output=True, text=True)
        hermione_planet_result = subprocess.run(hermione_planet, shell=True, check=True, capture_output=True, text=True)
        ixion_planet_result = subprocess.run(ixion_planet, shell=True, check=True, capture_output=True, text=True)
        haumea_planet_result = subprocess.run(haumea_planet, shell=True, check=True, capture_output=True, text=True)
        makemake_planet_result = subprocess.run(makemake_planet, shell=True, check=True, capture_output=True, text=True)
        bamberga_planet_result = subprocess.run(bamberga_planet, shell=True, check=True, capture_output=True, text=True)
        patientia_planet_result = subprocess.run(patientia_planet, shell=True, check=True, capture_output=True, text=True)
        thisbe_planet_result = subprocess.run(thisbe_planet, shell=True, check=True, capture_output=True, text=True)
        herculina_planet_result = subprocess.run(herculina_planet, shell=True, check=True, capture_output=True, text=True)
        doris_planet_result = subprocess.run(doris_planet, shell=True, check=True, capture_output=True, text=True)
        ursula_planet_result = subprocess.run(ursula_planet, shell=True, check=True, capture_output=True, text=True)
        eugenia_planet_result = subprocess.run(eugenia_planet, shell=True, check=True, capture_output=True, text=True)
        amphitrite_planet_result = subprocess.run(amphitrite_planet, shell=True, check=True, capture_output=True, text=True)
        diotima_planet_result = subprocess.run(diotima_planet, shell=True, check=True, capture_output=True, text=True)
        fortuna_planet_result = subprocess.run(fortuna_planet, shell=True, check=True, capture_output=True, text=True)
        egeria_planet_result = subprocess.run(egeria_planet, shell=True, check=True, capture_output=True, text=True)
        themis_planet_result = subprocess.run(themis_planet, shell=True, check=True, capture_output=True, text=True)
        aurora_planet_result = subprocess.run(aurora_planet, shell=True, check=True, capture_output=True, text=True)
        alauda_planet_result = subprocess.run(alauda_planet, shell=True, check=True, capture_output=True, text=True)
        aletheia_planet_result = subprocess.run(aletheia_planet, shell=True, check=True, capture_output=True, text=True)
        palma_planet_result = subprocess.run(palma_planet, shell=True, check=True, capture_output=True, text=True)
        nemesis_planet_result = subprocess.run(nemesis_planet, shell=True, check=True, capture_output=True, text=True)
        psyche_planet_result = subprocess.run(psyche_planet, shell=True, check=True, capture_output=True, text=True)
        hebe_planet_result =  subprocess.run(hebe_planet, shell=True, check=True, capture_output=True, text=True)
        lachesis_planet_result = subprocess.run(lachesis_planet, shell=True, check=True, capture_output=True, text=True)
        daphne_planet_result = subprocess.run(daphne_planet, shell=True, check=True, capture_output=True, text=True)
        bertha_planet_result = subprocess.run(bertha_planet, shell=True, check=True, capture_output=True, text=True)
        freia_planet_result = subprocess.run(freia_planet, shell=True, check=True, capture_output=True, text=True)
        winchester_planet_result = subprocess.run(winchester_planet, shell=True, check=True, capture_output=True, text=True)
        hilda_planet_result = subprocess.run(hilda_planet, shell=True, check=True, capture_output=True, text=True)
        pretoria_planet_result = subprocess.run(pretoria_planet, shell=True, check=True, capture_output=True, text=True)
        metis_planet_result = subprocess.run(metis_planet, shell=True, check=True, capture_output=True, text=True)
        aegle_planet_result = subprocess.run(aegle_planet, shell=True, check=True, capture_output=True, text=True)
        kalliope_planet_result = subprocess.run(kalliope_planet, shell=True, check=True, capture_output=True, text=True)
        germania_planet_result = subprocess.run(germania_planet, shell=True, check=True, capture_output=True, text=True)
        prokne_planet_result = subprocess.run(prokne_planet, shell=True, check=True, capture_output=True, text=True)
        stereoskopia_planet_result = subprocess.run(stereoskopia_planet, shell=True, check=True, capture_output=True, text=True)
        agamemnon_planet_result = subprocess.run(agamemnon_planet, shell=True, check=True, capture_output=True, text=True)
        alexandra_planet_result = subprocess.run(alexandra_planet, shell=True, check=True, capture_output=True, text=True)
        siegena_planet_result = subprocess.run(siegena_planet, shell=True, check=True, capture_output=True, text=True)
        elpis_planet_result = subprocess.run(elpis_planet, shell=True, check=True, capture_output=True, text=True)
        lilith_real_planet_result = subprocess.run(lilith_real_planet, shell=True, check=True, capture_output=True, text=True)
        borasisi_planet_result = subprocess.run(borasisi_planet, shell=True, check=True, capture_output=True, text=True)
        lempo_planet_result = subprocess.run(lempo_planet, shell=True, check=True, capture_output=True, text=True)
        _1998_26308_planet_result = subprocess.run(_1998_26308_planet, shell=True, check=True, capture_output=True, text=True)
        ceto_planet_result = subprocess.run(ceto_planet, shell=True, check=True, capture_output=True, text=True)
        teharonhiawako_planet_result = subprocess.run(teharonhiawako_planet, shell=True, check=True, capture_output=True, text=True)
        _2000_oj67_134860_planet_result = subprocess.run(_2000_oj67_134860_planet, shell=True, check=True, capture_output=True, text=True)
        elektra_planet_result = subprocess.run(elektra_planet, shell=True, check=True, capture_output=True, text=True)
        typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        chicago_planet_result = subprocess.run(chicago_planet, shell=True, check=True, capture_output=True, text=True)
        loreley_planet_result = subprocess.run(loreley_planet, shell=True, check=True, capture_output=True, text=True)
        diomedes_planet_result = subprocess.run(diomedes_planet, shell=True, check=True, capture_output=True, text=True)
        kreusa_planet_result = subprocess.run(kreusa_planet, shell=True, check=True, capture_output=True, text=True)
        gyptis_planet_result = subprocess.run(gyptis_planet, shell=True, check=True, capture_output=True, text=True)
        juewa_planet_result = subprocess.run(juewa_planet, shell=True, check=True, capture_output=True, text=True)
        eunike_planet_result = subprocess.run(eunike_planet, shell=True, check=True, capture_output=True, text=True)
        ino_planet_result = subprocess.run(ino_planet, shell=True, check=True, capture_output=True, text=True)
        ismene_planet_result = subprocess.run(ismene_planet, shell=True, check=True, capture_output=True, text=True)
        merapi_planet_result = subprocess.run(merapi_planet, shell=True, check=True, capture_output=True, text=True)

        astro_objects = [
    "Chiron",
    "Lilith",
    "Vertex",
    "Ceres",
    "Pallas",
    "Juno",
    "Vesta",
    "Eris",
    "Selena/White Moon",
    "Quaoar",
    "Sedna",
    "Varuna",
    "Nessus",
    "Waldemath",
    "Hygiea",
    "Sylvia",
    "Hektor",
    "Europa",
    "Davida",
    "Interamnia",
    "Camilla",
    "Cybele",
    "Black Sun",
    "Anti-Vertex",
    "True South Node",
    "True Black Sun",
    "Lilith 2",
    "Waldemath Priapus",
    "White Sun",
    "Chariklo",
    "Iris",
    "Eunomia",
    "Euphrosyne",
    "Orcus",
    "Pholus",
    "Hermione",
    "Ixion",
    "Haumea",
    "Makemake",
    "Bamberga",
    "Patientia",
    "Thisbe",
    "Herculina",
    "Doris",
    "Ursula",
    "Eugenia",
    "Amphitrite",
    "Diotima",
    "Fortuna",
    "Egeria",
    "Themis",
    "Aurora",
    "Alauda",
    "Aletheia",
    "Palma",
    "Nemesis",
    "Psyche",
    "Hebe",
    "Lachesis",
    "Daphne",
    "Bertha",
    "Freia",
    "Winchester",
    "Hilda",
    "Pretoria",
    "Metis",
    "Aegle",
    "Kalliope",
    "Germania",
    "Prokne",
    "Stereoskopia",
    "Agamemnon",
    "Alexandra",
    "Siegena",
    "Elpis",
    "Real Lilith",
    "Black Sun 2",
    "Vulcan",
    "Borasisi",
    "Lempo",
    "1998 SM165",
    "Ceto",
    "Teharonhiawako",
    "2000 OJ67",
    "Elektra",
    "Typhon",
    "Aspasia",
    "Chicago",
    "Loreley",
    "Gyptis",
    "Diomedes",
    "Kreusa",
    "Juewa",
    "Eunike",
    "Ino",
    "Ismene",
    "Merapi"
]

        houses_objects = [
            "house  1",
            "house  2",
            "house  3",
            "house  4",
            "house  5",
            "house  6",
            "Vertex"
        ]
        planets_object = [
            'Sun',
            'Moon',
            'Mercury',
            'Venus',
            'Mars',
            'Jupiter',
            'Saturn',
            'Uranus',
            'Neptune',
            'Pluto',
            'true Node'

        ]
        output = result.stdout
        # First House
        houses_1_parse_output = parse_houses_and_vertex(output,1)
        # # Second House 
        houses_2_parse_output = parse_houses_and_vertex(output,2)
        # # Third House
        houses_3_parse_output = parse_houses_and_vertex(output,3)
        # # Fourth House
        houses_4_parse_output = parse_houses_and_vertex(output,4)
        # # Fifth House
        houses_5_parse_output = parse_houses_and_vertex(output,5)
        # # Sixth House
        houses_6_parse_output = parse_houses_and_vertex(output,6)
        # # Vertex
        houses_vertex_parse_output = parse_houses_and_vertex(output,houses_objects[6])

        # print(houses_parse_output)
        print(f"Data of the Houses: {houses_1_parse_output}")
        


        output2 = result2.stdout
        lines2 = output2.splitlines()



        # # Output of the Planets 
        planet_sun_parse_output = parse_planets(output2,planets_object[0])
        print(f"Data of the Sun: {planet_sun_parse_output}")
        planet_moon_parse_output = parse_planets(output2,planets_object[1])
        print(f"Data of the Moon: {planet_moon_parse_output}")
        planet_mercury_parse_output = parse_planets(output2,planets_object[2])
        print(f"Data of the Mercury: {planet_mercury_parse_output}")
        planet_venus_parse_output = parse_planets(output2,planets_object[3])
        print(f"Data of the Venus: {planet_venus_parse_output}")
        planet_mars_parse_output = parse_planets(output2,planets_object[4])
        print(f"Data of the Mars: {planet_mars_parse_output}")
        planet_jupiter_parse_output = parse_planets(output2,planets_object[5])
        print(f"Data of the Jupiter: {planet_jupiter_parse_output}")
        planet_saturn_parse_output = parse_planets(output2,planets_object[6])
        print(f"Data of the Saturn: {planet_saturn_parse_output}")
        planet_uranus_parse_output = parse_planets(output2,planets_object[7])
        print(f"Data of the Uranus: {planet_uranus_parse_output}")
        planet_neptune_parse_output = parse_planets(output2,planets_object[8])
        print(f"Data of the Neptune: {planet_neptune_parse_output}")
        planet_pluto_parse_output = parse_planets(output2,planets_object[9])
        print(f"Data of the Pluto: {planet_pluto_parse_output}")
        planet_true_node_parse_output = parse_planets(output2,planets_object[10])
        print(f"Data of the True Node: {planet_true_node_parse_output}")
        # Hypothetical Planet 
        lilith_real_planet_result_output = lilith_real_planet_result.stdout
        lilith_real_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'osc. Apogee')
        # Sol Negro 2
        sol_blanco_planet_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'intp. Perigee')

        # Vulcan
        vulcan_planet_parse_output = parse_asteroid_output(lilith_real_planet_result_output,'Vulcan ')

        



        quiron_output = quiron_planet_result.stdout
        quiron_parse_output= parse_asteroid_output(quiron_output,astro_objects[0])
        
        lilith_output = lilith_planet_result.stdout
        lilith_parse_output = parse_asteroid_output(lilith_output,astro_objects[1])

        cerus_output = cerus_planet_result.stdout
        cerus_parse_output = parse_asteroid_output(cerus_output,astro_objects[3])

        pallas_output = pallas_planet_result.stdout
        pallas_parse_output = parse_asteroid_output(pallas_output,astro_objects[4])

        juno_output = juno_planet_result.stdout
        juno_parse_output = parse_asteroid_output(juno_output,astro_objects[5])

        vesta_output = vesta_planet_result.stdout
        vesta_parse_output = parse_asteroid_output(vesta_output,astro_objects[6])

        eris_output = eris_planet_result.stdout
        eris_parse_output = parse_asteroid_output(eris_output,astro_objects[7])

        white_moon_output = white_moon_result.stdout
        white_moon_parse_output = parse_asteroid_output(white_moon_output,astro_objects[8])

        quaoar_output = quaoar_planet_result.stdout
        quaoar_parse_output = parse_asteroid_output(quaoar_output,astro_objects[9])

        sedna_output = sedna_planet_result.stdout
        sedna_parse_output = parse_asteroid_output(sedna_output,astro_objects[10])

        varuna_output = varuna_planet_result.stdout
        varuna_parse_output = parse_asteroid_output(varuna_output,astro_objects[11])

        nessus_output = nessus_planet_result.stdout
        nessus_parse_output = parse_asteroid_output(nessus_output,astro_objects[12])

        waltemath_output = waltemath_planet_result.stdout
        waltemath_parse_output = parse_asteroid_output(waltemath_output,astro_objects[13])

        hygeia_output = hygeia_planet_result.stdout
        hygeia_parse_output = parse_asteroid_output(hygeia_output,astro_objects[14])

        sylvia_output = sylvia_planet_result.stdout
        sylvia_parse_output = parse_asteroid_output(sylvia_output,astro_objects[15])

        hektor_output = hektor_planet_result.stdout
        hektor_parse_output = parse_asteroid_output(hektor_output,astro_objects[16])

        europa_output = europa_planet_result.stdout
        europa_parse_output = parse_asteroid_output(europa_output,astro_objects[17])

        davida_output = davida_planet_result.stdout
        davida_parse_output = parse_asteroid_output(davida_output,astro_objects[18])

        interamnia_output = interamnia_planet_result.stdout
        interamnia_parse_output = parse_asteroid_output(interamnia_output,astro_objects[19])

        camilla_output = camilla_planet_result.stdout
        camilla_parse_output = parse_asteroid_output(camilla_output,astro_objects[20])

        cybele_output = cybele_planet_result.stdout
        cybele_parse_output = parse_asteroid_output(cybele_output,astro_objects[21])

        chariklo_output = chariklo_planet_result.stdout
        chariklo_parse_output = parse_asteroid_output(chariklo_output,astro_objects[29])

        iris_output = iris_planet_result.stdout
        iris_parse_output = parse_asteroid_output(iris_output,astro_objects[30])

        eunomia_planet_output = eunomia_planet_result.stdout
        eunomia_parse_output = parse_asteroid_output(eunomia_planet_output,astro_objects[31])

        euphrosyne_output = euphrosyne_planet_result.stdout
        euphrosyne_parse_output = parse_asteroid_output(euphrosyne_output,astro_objects[32])

        orcus_output = orcus_planet_result.stdout
        orcus_parse_output = parse_asteroid_output(orcus_output,astro_objects[33])

        pholus_output = pholus_planet_result.stdout
        pholus_parse_output = parse_asteroid_output(pholus_output,astro_objects[34])

        hermione_output = hermione_planet_result.stdout
        hermione_parse_output = parse_asteroid_output(hermione_output,astro_objects[35])

        ixion_output = ixion_planet_result.stdout
        ixion_parse_output = parse_asteroid_output(ixion_output,astro_objects[36])

        haumea_output = haumea_planet_result.stdout
        haumea_parse_output = parse_asteroid_output(haumea_output,astro_objects[37])

        makemake_output = makemake_planet_result.stdout
        makemake_parse_output = parse_asteroid_output(makemake_output,astro_objects[38])

        bamberga_output = bamberga_planet_result.stdout
        bamberga_parse_output = parse_asteroid_output(bamberga_output,astro_objects[39])

        patientia_output = patientia_planet_result.stdout
        patientia_parse_output = parse_asteroid_output(patientia_output,astro_objects[40])

        thisbe_output = thisbe_planet_result.stdout
        thisbe_parse_output = parse_asteroid_output(thisbe_output,astro_objects[41])

        herculina_output = herculina_planet_result.stdout
        herculina_parse_output = parse_asteroid_output(herculina_output,astro_objects[42])

        doris_output = doris_planet_result.stdout
        doris_parse_output = parse_asteroid_output(doris_output,astro_objects[43])  

        ursula_output = ursula_planet_result.stdout
        ursula_parse_output = parse_asteroid_output(ursula_output,astro_objects[44])

        eugenia_output = eugenia_planet_result.stdout
        eugenia_parse_output = parse_asteroid_output(eugenia_output,astro_objects[45])

        amphitrite_output = amphitrite_planet_result.stdout
        amphitrite_parse_output = parse_asteroid_output(amphitrite_output,astro_objects[46])

        diotima_output = diotima_planet_result.stdout
        diotima_parse_output = parse_asteroid_output(diotima_output,astro_objects[47])

        fortuna_output = fortuna_planet_result.stdout
        fortuna_parse_output = parse_asteroid_output(fortuna_output,astro_objects[48])

        egeria_output = egeria_planet_result.stdout
        egeria_parse_output = parse_asteroid_output(egeria_output,astro_objects[49])

        themis_output = themis_planet_result.stdout
        themis_parse_output = parse_asteroid_output(themis_output,astro_objects[50])

        aurora_output = aurora_planet_result.stdout
        aurora_parse_output = parse_asteroid_output(aurora_output,astro_objects[51])

        alauda_output = alauda_planet_result.stdout
        alauda_parse_output = parse_asteroid_output(alauda_output,astro_objects[52])

        aletheia_output = aletheia_planet_result.stdout
        aletheia_parse_output = parse_asteroid_output(aletheia_output,astro_objects[53])

        palma_output = palma_planet_result.stdout
        palma_parse_output = parse_asteroid_output(palma_output,astro_objects[54])

        nemesis_output = nemesis_planet_result.stdout
        nemesis_parse_output = parse_asteroid_output(nemesis_output,astro_objects[55])

        psyche_output = psyche_planet_result.stdout
        psyche_parse_output = parse_asteroid_output(psyche_output,astro_objects[56])

        hebe_output = hebe_planet_result.stdout
        hebe_parse_output = parse_asteroid_output(hebe_output,astro_objects[57])

        lachesis_output = lachesis_planet_result.stdout
        lachesis_parse_output = parse_asteroid_output(lachesis_output,astro_objects[58])

        daphne_output = daphne_planet_result.stdout
        daphne_parse_output = parse_asteroid_output(daphne_output,astro_objects[59])

        bertha_output = bertha_planet_result.stdout
        bertha_parse_output = parse_asteroid_output(bertha_output,astro_objects[60])

        freia_output = freia_planet_result.stdout
        freia_parse_output = parse_asteroid_output(freia_output,astro_objects[61])

        winchester_output = winchester_planet_result.stdout
        winchester_parse_output = parse_asteroid_output(winchester_output,astro_objects[62])
        

        hilda_output = hilda_planet_result.stdout
        hilda_parse_output = parse_asteroid_output(hilda_output,astro_objects[63])

        pretoria_output = pretoria_planet_result.stdout
        pretoria_parse_output = parse_asteroid_output(pretoria_output,astro_objects[64])

        metis_output = metis_planet_result.stdout
        metis_parse_output = parse_asteroid_output(metis_output,astro_objects[65])

        aegle_output = aegle_planet_result.stdout
        aegle_parse_output = parse_asteroid_output(aegle_output,astro_objects[66])

        kalliope_output = kalliope_planet_result.stdout
        kalliope_parse_output = parse_asteroid_output(kalliope_output,astro_objects[67])

        germania_output = germania_planet_result.stdout
        germania_parse_output = parse_asteroid_output(germania_output,astro_objects[68])

        prokne_output = prokne_planet_result.stdout
        prokne_parse_output = parse_asteroid_output(prokne_output,astro_objects[69])

        stereoskopia_output = stereoskopia_planet_result.stdout
        stereoskopia_parse_output = parse_asteroid_output(stereoskopia_output,astro_objects[70])

        agamemnon_output = agamemnon_planet_result.stdout
        agamemnon_parse_output = parse_asteroid_output(agamemnon_output,astro_objects[71])

        alexandra_output = alexandra_planet_result.stdout
        alexandra_parse_output = parse_asteroid_output(alexandra_output,astro_objects[72])

        siegena_output = siegena_planet_result.stdout
        siegena_parse_output = parse_asteroid_output(siegena_output,astro_objects[73])

        elpis_output = elpis_planet_result.stdout
        elpis_parse_output = parse_asteroid_output(elpis_output,astro_objects[74])

        borasisi_output = borasisi_planet_result.stdout
        borasisi_parse_output = parse_asteroid_output(borasisi_output,astro_objects[78])

        lempo_output = lempo_planet_result.stdout
        lempo_parse_output = parse_asteroid_output(lempo_output,astro_objects[79])

        _1998_26308_output = _1998_26308_planet_result.stdout
        _1998_26308_parse_output = parse_asteroid_output(_1998_26308_output,astro_objects[80])

        ceto_output = ceto_planet_result.stdout
        ceto_parse_output = parse_asteroid_output(ceto_output,astro_objects[81])

        teharonhiawako_output = teharonhiawako_planet_result.stdout
        teharonhiawako_parse_output = parse_asteroid_output(teharonhiawako_output,astro_objects[82])

        _2000_oj67_134860_output = _2000_oj67_134860_planet_result.stdout
        _2000_oj67_134860_parse_output = parse_asteroid_output(_2000_oj67_134860_output,astro_objects[83])

        elektra_output = elektra_planet_result.stdout
        elektra_parse_output = parse_asteroid_output(elektra_output,astro_objects[84])

        typhon_output = typhon_planet_result.stdout
        typhon_parse_output = parse_asteroid_output(typhon_output,astro_objects[85])

        aspasia_output = aspasia_planet_result.stdout
        aspasia_parse_output = parse_asteroid_output(aspasia_output,astro_objects[86])

        chicago_output = chicago_planet_result.stdout
        chicago_parse_output = parse_asteroid_output(chicago_output,astro_objects[87])

        loreley_output = loreley_planet_result.stdout
        loreley_parse_output = parse_asteroid_output(loreley_output,astro_objects[88])

        gyptis_output = gyptis_planet_result.stdout
        gyptis_parse_output = parse_asteroid_output(gyptis_output,astro_objects[89])

        diomedes_output = diomedes_planet_result.stdout
        diomedes_parse_output = parse_asteroid_output(diomedes_output,astro_objects[90])

 

        kreusa_output = kreusa_planet_result.stdout
        kreusa_parse_output = parse_asteroid_output(kreusa_output,astro_objects[91])

        juewa_output = juewa_planet_result.stdout
        juewa_parse_output = parse_asteroid_output(juewa_output,astro_objects[92])

        eunike_output = eunike_planet_result.stdout
        eunike_parse_output = parse_asteroid_output(eunike_output,astro_objects[93])

        ino_output = ino_planet_result.stdout
        ino_parse_output = parse_asteroid_output(ino_output,astro_objects[94])

        ismene_output = ismene_planet_result.stdout
        ismene_parse_output = parse_asteroid_output(ismene_output,astro_objects[95])

        merapi_output = merapi_planet_result.stdout
        merapi_parse_output = parse_asteroid_output(merapi_output,astro_objects[96])



        # Create a dictionary to store the result data that are Empty in the Excel
        sol_negro_parse_output =  {
            "name": "Sol Negro",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
            "retrograde": ""
        }
        # For AntiVertex
        anti_vertex_parse_output = {
             "name": "Antivertex",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "" ,
             "retrograde": ""
        }
        # For Nodo Sur Real
        nodo_sur_real_parse_output = {
               "name": "Nodo Sur Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""  
        }
        # For Sol Negro Real
        sol_negro_real_parse_output = {
            "name": "Sol Negro Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # For Lilith 2
        lilith2_parse_output = {
            "name": "Lilith 2",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # For Waldemath Priapus
        waltemath_priapus_parse_output = {
            "name": "Waldemath Priapus",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        }
        # Sol Blanco 
        sol_blanco_parse_output = {
            "name": "Sol Blanco",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": "",
             "retrograde": ""
        } 

        planets = []
                # List of Data for Natal Positions
        asteroidsList = [houses_1_parse_output,houses_2_parse_output,houses_3_parse_output,houses_4_parse_output,houses_5_parse_output,houses_6_parse_output,planet_sun_parse_output,planet_moon_parse_output,planet_mercury_parse_output,planet_venus_parse_output,planet_mars_parse_output,planet_jupiter_parse_output,planet_saturn_parse_output,planet_uranus_parse_output,planet_neptune_parse_output,planet_pluto_parse_output,planet_true_node_parse_output,quiron_parse_output,lilith_parse_output,houses_vertex_parse_output,cerus_parse_output,pallas_parse_output,juno_parse_output,vesta_parse_output,eris_parse_output,white_moon_parse_output,quaoar_parse_output,sedna_parse_output,varuna_parse_output,nessus_parse_output,waltemath_parse_output,hygeia_parse_output,sylvia_parse_output,hektor_parse_output,europa_parse_output,davida_parse_output,interamnia_parse_output,camilla_parse_output,cybele_parse_output,sol_negro_parse_output,anti_vertex_parse_output,nodo_sur_real_parse_output,sol_negro_real_parse_output,lilith2_parse_output,waltemath_priapus_parse_output,sol_blanco_parse_output,chariklo_parse_output,iris_parse_output,eunomia_parse_output,euphrosyne_parse_output,orcus_parse_output,pholus_parse_output,hermione_parse_output,ixion_parse_output,haumea_parse_output,makemake_parse_output,bamberga_parse_output,patientia_parse_output,thisbe_parse_output,herculina_parse_output,doris_parse_output,ursula_parse_output,eugenia_parse_output,amphitrite_parse_output,diotima_parse_output,fortuna_parse_output,egeria_parse_output,themis_parse_output,aurora_parse_output,alauda_parse_output,aletheia_parse_output,palma_parse_output,nemesis_parse_output,psyche_parse_output,hebe_parse_output,lachesis_parse_output,daphne_parse_output,bertha_parse_output,freia_parse_output,winchester_parse_output,hilda_parse_output,pretoria_parse_output,metis_parse_output,aegle_parse_output,kalliope_parse_output,germania_parse_output,prokne_parse_output,stereoskopia_parse_output,agamemnon_parse_output,alexandra_parse_output,siegena_parse_output,elpis_parse_output,lilith_real_parse_output,sol_blanco_planet_parse_output,vulcan_planet_parse_output,borasisi_parse_output,lempo_parse_output,_1998_26308_parse_output,ceto_parse_output,teharonhiawako_parse_output,_2000_oj67_134860_parse_output,elektra_parse_output,typhon_parse_output,aspasia_parse_output,chicago_parse_output,loreley_parse_output,gyptis_parse_output,diomedes_parse_output,kreusa_parse_output,juewa_parse_output,eunike_parse_output,ino_parse_output,ismene_parse_output,merapi_parse_output]
    
        return asteroidsList
   
   
    except Exception as e:
        print("Error initializing Excel:", e)
        logger.error(f"Error occurred: {str(e)}\n{traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library


def close_excel_without_save():
    # Create an instance of the Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    
    # Optional: Set this to True if you want to make Excel visible while running the script
    excel.Visible = False
    
    # Prevent the "Do you want to save changes?" prompt
    excel.DisplayAlerts = False
    
    # Loop through all the open workbooks
    for wb in excel.Workbooks:
        # Mark each workbook as saved, so Excel won't ask to save
        wb.Saved = True
    
    # Quit the Excel application without saving any changes
    excel.Quit()
