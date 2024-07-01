from flask import Blueprint, jsonify, request
import subprocess
import re
import win32com.client
import pythoncom
import os
import shutil
from datetime import datetime

change_excel_general = Blueprint('change_excel_general', __name__)

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

@change_excel_general.route('/change_excel_general', methods=['POST'])
def run_excel_macro_changeData():
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

        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False  # Set to True if you want Excel to be visible

        # Construct the command with zero-padded values
        # For House Data From Cell D5 to D10
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
        # For Planets Data From Cell D11 to D21 Which Includes True Node
        command2 = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZ -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"
        # For Quirón Command From Cell D22
        quiron_planet = f"swetest -ps -xs2060 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # For Lilith Command From Cell D23
        lilith_planet = f"swetest -ps -xs1181 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # For Cerus Command 
        cerus_planet = f"swetest -ps -xs1 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # For Pallas Command
        pallas_planet = f"swetest -ps -xs2 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # For Juno Command
        juno_planet = f"swetest -ps -xs3 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # For Vesta Command
        vesta_planet = f"swetest -ps -xs4 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Eris Command
        eris_planet = f"swetest -ps -xs136199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # White Moon Command
        white_moon = f"swetest -pZ -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Quaoar Command
        quaoar_planet = f"swetest -ps -xs50000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Sedna Command
        sedna_planet = f"swetest -ps -xs90377 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Varuna Command
        varuna_planet = f"swetest -ps -xs20000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Nessus Command
        nessus_planet = f"swetest -ps -xs7066 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Waltemath Command
        waltemath_planet = f"swetest -pw -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Hygeia Command
        hygeia_planet = f"swetest -ps -xs10 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Sylvia Command
        sylvia_planet = f"swetest -ps -xs87 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
#         Hektor	624	Hector
        hektor_planet = f"swetest -ps -xs624 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Europa	52	Europa
        europa_planet = f"swetest -ps -xs52 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Davida	511	Davida
        davida_planet = f"swetest -ps -xs511 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Interamnia	704	Interamnia
        interamnia_planet = f"swetest -ps -xs704 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Camilla	107	Camilla
        camilla_planet = f"swetest -ps -xs107 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Cybele	65	Cybele
        cybele_planet = f"swetest -ps -xs65 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Sol Negro	h22	Black Sun

# Antivertex		Anti-Vertex

# Nodo Sur Real		True South Node
# Sol Negro Real		True Black Sun
# Lilith 2		Lilith 2
# Waldemath Priapus		Waldemath Priapus
# Sol Blanco		White Sun
# Chariklo	10199	Chariklo
        chariklo_planet = f"swetest -ps -xs10199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Iris	7	Iris
        iris_planet = f"swetest -ps -xs7 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Eunomia	15	Eunomia
        eunomia_planet = f"swetest -ps -xs15 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Euphrosyne	31	Euphrosyne
        euphrosyne_planet = f"swetest -ps -xs31 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Orcus	90482	Orcus
        # Orcus Command
        orcus_planet = f"swetest -ps -xs90482 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Pholus	5145	Pholus
        # Pholus Command
        pholus_planet = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Hermíone Command
        hermione_planet = f"swetest -ps -xs121 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Ixion Command
        ixion_planet = f"swetest -ps -xs28978 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Haumea Command
        haumea_planet = f"swetest -ps -xs136108 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Makemake Command
        makemake_planet = f"swetest -ps -xs136472 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Bamberga Command
        bamberga_planet = f"swetest -ps -xs324 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Patientia Command
        patientia_planet = f"swetest -ps -xs451 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Thisbe Command
        thisbe_planet = f"swetest -ps -xs88 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Herculina Command
        herculina_planet = f"swetest -ps -xs532 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Doris Command
        doris_planet = f"swetest -ps -xs48 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Ursula Command
        ursula_planet = f"swetest -ps -xs375 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Eugenia Command
        eugenia_planet = f"swetest -ps -xs45 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Amphitrite Command
        amphitrite_planet = f"swetest -ps -xs29 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Diotima Command
        diotima_planet = f"swetest -ps -xs423 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Fortuna Command
        fortuna_planet = f"swetest -ps -xs19 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Egeria Command
        egeria_planet = f"swetest -ps -xs13 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Themis Command
        themis_planet = f"swetest -ps -xs24 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Aurora Command
        aurora_planet = f"swetest -ps -xs94 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Alauda Command
        alauda_planet = f"swetest -ps -xs702 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Aletheia Command
        aletheia_planet = f"swetest -ps -xs259 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Palma Command
        palma_planet = f"swetest -ps -xs372 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Nemesis Command
        nemesis_planet = f"swetest -ps -xs128 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Psyche Command
        psyche_planet = f"swetest -ps -xs16 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Hebe Command
        hebe_planet = f"swetest -ps -xs6 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Lachesis Command
        lachesis_planet = f"swetest -ps -xs120 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Daphne Command
        daphne_planet = f"swetest -ps -xs41 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Bertha Command
        bertha_planet = f"swetest -ps -xs154 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Freia Command
        freia_planet = f"swetest -ps -xs76 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Winchester Command
        winchester_planet = f"swetest -ps -xs747 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Hilda Command
        hilda_planet = f"swetest -ps -xs153 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Pretoria Command
        pretoria_planet = f"swetest -ps -xs790 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Metis Command
        metis_planet = f"swetest -ps -xs9 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Aegle Command
        aegle_planet = f"swetest -ps -xs96 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Kalliope Command
        kalliope_planet = f"swetest -ps -xs22 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Germania Command
        germania_planet = f"swetest -ps -xs241 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Prokne Command
        prokne_planet = f"swetest -ps -xs194 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Stereoskopia Command
        stereoskopia_planet = f"swetest -ps -xs566 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Agamemnon Command
        agamemnon_planet = f"swetest -ps -xs911 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Alexandra Command
        alexandra_planet = f"swetest -ps -xs54 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Siegena Command
        siegena_planet = f"swetest -ps -xs386 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Elpis Command
        elpis_planet = f"swetest -ps -xs59 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Lilith Real Command
        # lilith_real_planet = f"swetest -ps -xsh13 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Sol Negro 2 Command
            # TODO: Implementation of the Sol Negro 2 Planet
        # sol_negro_2_planet = f"swetest -ps -xsh22 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Vulcan Command
            # TODO: Implementation of the Vulcan Planet
        # vulcan_planet = f"swetest -ps -xsh55 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Borasisi Command
        borasisi_planet = f"swetest -ps -xs66652 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Lempo Command
        lempo_planet = f"swetest -ps -xs47171 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # 1998(26308) Command
        _1998_26308_planet = f"swetest -ps -xs26308 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Ceto Command
        ceto_planet = f"swetest -ps -xs65489 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Teharonhiawako Command
        teharonhiawako_planet = f"swetest -ps -xs88611 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # 2000 OJ67 (134860) Command
        _2000_oj67_134860_planet = f"swetest -ps -xs134860 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Elektra Command
        elektra_planet = f"swetest -ps -xs130 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Typhon Command
        typhon_planet = f"swetest -ps -xs42355 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Aspasia Command
        aspasia_planet = f"swetest -ps -xs409 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Chicago Command
        chicago_planet = f"swetest -ps -xs334 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Loreley Command
        loreley_planet = f"swetest -ps -xs165 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Gyptis Command
        gyptis_planet = f"swetest -ps -xs444 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Diomedes Command
        diomedes_planet = f"swetest -ps -xs1437 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # Kreusa Command
        kreusa_planet = f"swetest -ps -xs105 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Juewa	139	Juewa
        juewa_planet = f"swetest -ps -xs139 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Eunike	185	Eunike
        eunike_planet = f"swetest -ps -xs185 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Ino	173	Ino
        ino_planet = f"swetest -ps -xs173 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Ismene	190	Ismene
        ismene_planet = f"swetest -ps -xs190 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
# Merapi	536	Merapi
        merapi_planet = f"swetest -ps -xs536 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"


        # Execute the command using subprocess
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
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
        gyptis_planet_result = subprocess.run(gyptis_planet, shell=True, check=True, capture_output=True, text=True)
        juewa_planet_result = subprocess.run(juewa_planet, shell=True, check=True, capture_output=True, text=True)
        eunike_planet_result = subprocess.run(eunike_planet, shell=True, check=True, capture_output=True, text=True)
        ino_planet_result = subprocess.run(ino_planet, shell=True, check=True, capture_output=True, text=True)
        ismene_planet_result = subprocess.run(ismene_planet, shell=True, check=True, capture_output=True, text=True)
        merapi_planet_result = subprocess.run(merapi_planet, shell=True, check=True, capture_output=True, text=True)

        
        









        output = result.stdout
        lines = output.splitlines()

        output2 = result2.stdout
        lines2 = output2.splitlines()

        quiron_output = quiron_planet_result.stdout
        quiron_parse_output= parse_asteroid_output(quiron_output)
        
        lilith_output = lilith_planet_result.stdout
        lilith_parse_output = parse_asteroid_output(lilith_output)

        cerus_output = cerus_planet_result.stdout
        cerus_parse_output = parse_asteroid_output(cerus_output)

        pallas_output = pallas_planet_result.stdout
        pallas_parse_output = parse_asteroid_output(pallas_output)

        juno_output = juno_planet_result.stdout
        juno_parse_output = parse_asteroid_output(juno_output)

        vesta_output = vesta_planet_result.stdout
        vesta_parse_output = parse_asteroid_output(vesta_output)

        eris_output = eris_planet_result.stdout
        eris_parse_output = parse_asteroid_output(eris_output)

        white_moon_output = white_moon_result.stdout
        # white_moon_parse_output = parse_asteroid_output(white_moon_output)

        quaoar_output = quaoar_planet_result.stdout
        quaoar_parse_output = parse_asteroid_output(quaoar_output)

        sedna_output = sedna_planet_result.stdout
        sedna_parse_output = parse_asteroid_output(sedna_output)

        varuna_output = varuna_planet_result.stdout
        varuna_parse_output = parse_asteroid_output(varuna_output)

        nessus_output = nessus_planet_result.stdout
        nessus_parse_output = parse_asteroid_output(nessus_output)

        waltemath_output = waltemath_planet_result.stdout
        waltemath_parse_output = parse_asteroid_output(waltemath_output)

        hygeia_output = hygeia_planet_result.stdout
        hygeia_parse_output = parse_asteroid_output(hygeia_output)

        sylvia_output = sylvia_planet_result.stdout
        sylvia_parse_output = parse_asteroid_output(sylvia_output)

        hektor_output = hektor_planet_result.stdout
        hektor_parse_output = parse_asteroid_output(hektor_output)

        europa_output = europa_planet_result.stdout
        europa_parse_output = parse_asteroid_output(europa_output)

        davida_output = davida_planet_result.stdout
        davida_parse_output = parse_asteroid_output(davida_output)

        interamnia_output = interamnia_planet_result.stdout
        interamnia_parse_output = parse_asteroid_output(interamnia_output)

        camilla_output = camilla_planet_result.stdout
        camilla_parse_output = parse_asteroid_output(camilla_output)

        cybele_output = cybele_planet_result.stdout
        cybele_parse_output = parse_asteroid_output(cybele_output)

        chariklo_output = chariklo_planet_result.stdout
        chariklo_parse_output = parse_asteroid_output(chariklo_output)

        iris_output = iris_planet_result.stdout
        iris_parse_output = parse_asteroid_output(iris_output)

        eunomia_planet_output = eunomia_planet_result.stdout
        eunomia_parse_output = parse_asteroid_output(eunomia_planet_output)

        # TODO: Implement the parse_asteroid_output function for the following planets Parsing is not Correct Because Before and After the Degree The value will be 1 or 2 and don't forger the points
        euphrosyne_output = euphrosyne_planet_result.stdout
        # euphrosyne_parse_output = parse_asteroid_output(euphrosyne_output)

        orcus_output = orcus_planet_result.stdout
        orcus_parse_output = parse_asteroid_output(orcus_output)

        pholus_output = pholus_planet_result.stdout
        pholus_parse_output = parse_asteroid_output(pholus_output)

        hermione_output = hermione_planet_result.stdout
        hermione_parse_output = parse_asteroid_output(hermione_output)

        ixion_output = ixion_planet_result.stdout
        ixion_parse_output = parse_asteroid_output(ixion_output)

        haumea_output = haumea_planet_result.stdout
        haumea_parse_output = parse_asteroid_output(haumea_output)

        makemake_output = makemake_planet_result.stdout
        makemake_parse_output = parse_asteroid_output(makemake_output)

        bamberga_output = bamberga_planet_result.stdout
        bamberga_parse_output = parse_asteroid_output(bamberga_output)

        patientia_output = patientia_planet_result.stdout
        patientia_parse_output = parse_asteroid_output(patientia_output)

        thisbe_output = thisbe_planet_result.stdout
        thisbe_parse_output = parse_asteroid_output(thisbe_output)

        herculina_output = herculina_planet_result.stdout
        herculina_parse_output = parse_asteroid_output(herculina_output)

        doris_output = doris_planet_result.stdout
        doris_parse_output = parse_asteroid_output(doris_output)

        ursula_output = ursula_planet_result.stdout
        ursula_parse_output = parse_asteroid_output(ursula_output)

        eugenia_output = eugenia_planet_result.stdout
        eugenia_parse_output = parse_asteroid_output(eugenia_output)

        amphitrite_output = amphitrite_planet_result.stdout
        amphitrite_parse_output = parse_asteroid_output(amphitrite_output)

        diotima_output = diotima_planet_result.stdout
        diotima_parse_output = parse_asteroid_output(diotima_output)

        fortuna_output = fortuna_planet_result.stdout
        fortuna_parse_output = parse_asteroid_output(fortuna_output)

        egeria_output = egeria_planet_result.stdout
        egeria_parse_output = parse_asteroid_output(egeria_output)

        themis_output = themis_planet_result.stdout
        themis_parse_output = parse_asteroid_output(themis_output)

        aurora_output = aurora_planet_result.stdout
        aurora_parse_output = parse_asteroid_output(aurora_output)

        alauda_output = alauda_planet_result.stdout
        alauda_parse_output = parse_asteroid_output(alauda_output)

        aletheia_output = aletheia_planet_result.stdout
        aletheia_parse_output = parse_asteroid_output(aletheia_output)

        palma_output = palma_planet_result.stdout
        palma_parse_output = parse_asteroid_output(palma_output)

        nemesis_output = nemesis_planet_result.stdout
        nemesis_parse_output = parse_asteroid_output(nemesis_output)

        psyche_output = psyche_planet_result.stdout
        psyche_parse_output = parse_asteroid_output(psyche_output)

        hebe_output = hebe_planet_result.stdout
        hebe_parse_output = parse_asteroid_output(hebe_output)

        lachesis_output = lachesis_planet_result.stdout
        lachesis_parse_output = parse_asteroid_output(lachesis_output)

        daphne_output = daphne_planet_result.stdout
        daphne_parse_output = parse_asteroid_output(daphne_output)

        bertha_output = bertha_planet_result.stdout
        bertha_parse_output = parse_asteroid_output(bertha_output)

        freia_output = freia_planet_result.stdout
        freia_parse_output = parse_asteroid_output(freia_output)

        winchester_output = winchester_planet_result.stdout
        winchester_parse_output = parse_asteroid_output(winchester_output)

        hilda_output = hilda_planet_result.stdout
        hilda_parse_output = parse_asteroid_output(hilda_output)

        pretoria_output = pretoria_planet_result.stdout
        pretoria_parse_output = parse_asteroid_output(pretoria_output)

        metis_output = metis_planet_result.stdout
        metis_parse_output = parse_asteroid_output(metis_output)

        aegle_output = aegle_planet_result.stdout
        aegle_parse_output = parse_asteroid_output(aegle_output)

        kalliope_output = kalliope_planet_result.stdout
        kalliope_parse_output = parse_asteroid_output(kalliope_output)

        germania_output = germania_planet_result.stdout
        germania_parse_output = parse_asteroid_output(germania_output)

        prokne_output = prokne_planet_result.stdout
        prokne_parse_output = parse_asteroid_output(prokne_output)

        stereoskopia_output = stereoskopia_planet_result.stdout
        stereoskopia_parse_output = parse_asteroid_output(stereoskopia_output)

        #     agamemnon_planet_result = subprocess.run(agamemnon_planet, shell=True, check=True, capture_output=True, text=True)
        # alexandra_planet_result = subprocess.run(alexandra_planet, shell=True, check=True, capture_output=True, text=True)
        # siegena_planet_result = subprocess.run(siegena_planet, shell=True, check=True, capture_output=True, text=True)
        # elpis_planet_result = subprocess.run(elpis_planet, shell=True, check=True, capture_output=True, text=True)
        # borasisi_planet_result = subprocess.run(borasisi_planet, shell=True, check=True, capture_output=True, text=True)
        # lempo_planet_result = subprocess.run(lempo_planet, shell=True, check=True, capture_output=True, text=True)
        # _1998_26308_planet_result = subprocess.run(_1998_26308_planet, shell=True, check=True, capture_output=True, text=True)
        # ceto_planet_result = subprocess.run(ceto_planet, shell=True, check=True, capture_output=True, text=True)
        # teharonhiawako_planet_result = subprocess.run(teharonhiawako_planet, shell=True, check=True, capture_output=True, text=True)
        # _2000_oj67_134860_planet_result = subprocess.run(_2000_oj67_134860_planet, shell=True, check=True, capture_output=True, text=True)
        # elektra_planet_result = subprocess.run(elektra_planet, shell=True, check=True, capture_output=True, text=True)
        # typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        # aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        # typhon_planet_result = subprocess.run(typhon_planet, shell=True, check=True, capture_output=True, text=True)
        # aspasia_planet_result = subprocess.run(aspasia_planet, shell=True, check=True, capture_output=True, text=True)
        # chicago_planet_result = subprocess.run(chicago_planet, shell=True, check=True, capture_output=True, text=True)
        # gyptis_planet_result = subprocess.run(gyptis_planet, shell=True, check=True, capture_output=True, text=True)
        # diomedes_planet_result = subprocess.run(diomedes_planet, shell=True, check=True, capture_output=True, text=True)
        # gyptis_planet_result = subprocess.run(gyptis_planet, shell=True, check=True, capture_output=True, text=True)
        # juewa_planet_result = subprocess.run(juewa_planet, shell=True, check=True, capture_output=True, text=True)
        # eunike_planet_result = subprocess.run(eunike_planet, shell=True, check=True, capture_output=True, text=True)
        # ino_planet_result = subprocess.run(ino_planet, shell=True, check=True, capture_output=True, text=True)
        # ismene_planet_result = subprocess.run(ismene_planet, shell=True, check=True, capture_output=True, text=True)
        # merapi_planet_result = subprocess.run(merapi_planet, shell=True, check=True, capture_output=True, text=True)

        agamemnon_output = agamemnon_planet_result.stdout
        agamemnon_parse_output = parse_asteroid_output(agamemnon_output)

        alexandra_output = alexandra_planet_result.stdout
        alexandra_parse_output = parse_asteroid_output(alexandra_output)

        siegena_output = siegena_planet_result.stdout
        siegena_parse_output = parse_asteroid_output(siegena_output)

        elpis_output = elpis_planet_result.stdout
        elpis_parse_output = parse_asteroid_output(elpis_output)

        borasisi_output = borasisi_planet_result.stdout
        borasisi_parse_output = parse_asteroid_output(borasisi_output)

        lempo_output = lempo_planet_result.stdout
        lempo_parse_output = parse_asteroid_output(lempo_output)

        _1998_26308_output = _1998_26308_planet_result.stdout
        _1998_26308_parse_output = parse_asteroid_output(_1998_26308_output)

        ceto_output = ceto_planet_result.stdout
        ceto_parse_output = parse_asteroid_output(ceto_output)

        teharonhiawako_output = teharonhiawako_planet_result.stdout
        # teharonhiawako_parse_output = parse_asteroid_output(teharonhiawako_output)

        _2000_oj67_134860_output = _2000_oj67_134860_planet_result.stdout
        _2000_oj67_134860_parse_output = parse_asteroid_output(_2000_oj67_134860_output)

        elektra_output = elektra_planet_result.stdout
        elektra_parse_output = parse_asteroid_output(elektra_output)

        typhon_output = typhon_planet_result.stdout
        typhon_parse_output = parse_asteroid_output(typhon_output)

        aspasia_output = aspasia_planet_result.stdout
        aspasia_parse_output = parse_asteroid_output(aspasia_output)

        chicago_output = chicago_planet_result.stdout
        chicago_parse_output = parse_asteroid_output(chicago_output)

        gyptis_output = gyptis_planet_result.stdout
        gyptis_parse_output = parse_asteroid_output(gyptis_output)

        diomedes_output = diomedes_planet_result.stdout
        diomedes_parse_output = parse_asteroid_output(diomedes_output)

        loreley_output = loreley_planet_result.stdout
        loreley_parse_output = parse_asteroid_output(loreley_output)



        juewa_output = juewa_planet_result.stdout
        juewa_parse_output = parse_asteroid_output(juewa_output)

        eunike_output = eunike_planet_result.stdout
        eunike_parse_output = parse_asteroid_output(eunike_output)

        ino_output = ino_planet_result.stdout
        ino_parse_output = parse_asteroid_output(ino_output)

        ismene_output = ismene_planet_result.stdout
        ismene_parse_output = parse_asteroid_output(ismene_output)

        merapi_output = merapi_planet_result.stdout
        merapi_parse_output = parse_asteroid_output(merapi_output)



        # Create a dictionary to store the result data that are Empty in the Excel
        sol_negro_parse_output =  {
            "name": "Sol Negro",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""
        }
        # For AntiVertex
        anti_vertex_parse_output = {
             "name": "Antivertex",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""  
        }
        # For Nodo Sur Real
        nodo_sur_real_parse_output = {
               "name": "Nodo Sur Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""  
        }
        # For Sol Negro Real
        sol_negro_real_parse_output = {
            "name": "Sol Negro Real",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""
        }
        # For Lilith 2
        lilith2_parse_output = {
            "name": "Lilith 2",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""
        }
        # For Waldemath Priapus
        waltemath_priapus_parse_output = {
            "name": "Waldemath Priapus",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""
        }
        # Sol Blanco 
        sol_blanco_parse_output = {
            "name": "Sol Blanco",
            "positionDegree": "",
            "position_min": "",
            "position_sec": "",
            "position_sign": ""
        } 
        result_data = {}
        planets = []
        result_vertex = {}

        
        # Parse the output for houses
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 3 or more spaces
            for i in range(8, 14):  # Loop through lines 8 to 13 (houses 1 to 6)
                try:
                    match = re.split(pattern, lines[i])[1]
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', match)
                    degree_sign_abbr = degree_match_sign[0] if degree_match_sign else ""
                    degree_sign = zodiac_signs.get(degree_sign_abbr.lower(), degree_sign_abbr)
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    result_data[f"Casa{i - 7}"] = {
                        "positionDegree": int(degree_match.group(1)) if degree_match else None,
                        "position_sign": degree_sign,
                        "position_min": degree_match_min[0],
                        "position_sec": degree_match_min[1].replace('"', ''),  # Remove double quotes from seconds
                    }
                except IndexError as e:
                    result_data["error"] = f"Error parsing output: {str(e)}"
                    break
        else:
            result_data["error"] = "Error parsing line: No lines in the output"
        # Parse the output for houses
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 3 or more spaces
            for i in range(23, 24):  # Loop through lines 8 to 13 (houses 1 to 6)
                try:
                    match = re.split(pattern, lines[i])[1]
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', match)
                    degree_sign_abbr = degree_match_sign[0] if degree_match_sign else ""
                    degree_sign = zodiac_signs.get(degree_sign_abbr.lower(), degree_sign_abbr)
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    result_vertex = {
                        "name": re.split(pattern, lines[i])[0],
                        "positionDegree": int(degree_match.group(1)) if degree_match else None,
                        "position_sign": degree_sign,
                        "position_min": degree_match_min[0],
                        "position_sec": degree_match_min[1].replace('"', ''),  # Remove double quotes from seconds
                    }
                    
                except IndexError as e:
                    result_data["error"] = f"Error parsing output: {str(e)}"
                    break
        else:
            result_data["error"] = "Error parsing line: No lines in the output"
        
        # Parse the output for planets
        if len(lines2) > 0:
            planet_positions = lines2[6:16]
            planet_positions2 = lines2[17:18]

            for line in planet_positions:
                match = re.match(r"(\w+)\s+(.+)", line)
                if match:
                    planet_name = match.group(1)
                    position = match.group(2).strip()
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", position)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', position)
                    degree_sign_abbr = degree_match_sign[0] if degree_match_sign else ""
                    degree_sign = zodiac_signs.get(degree_sign_abbr.lower(), degree_sign_abbr)
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', position)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    
                    if degree_match:
                        degree = int(degree_match.group(1))
                        minute = degree_match_min[0]
                        second = degree_match_min[1]
                        
                        planets.append({
                            "planet_name": planet_name,
                            "positionDegree": degree,
                            "position_sign": degree_sign,
                            "position_min": minute,
                            "position_sec": second,
                        })

            for line in planet_positions2:
                pattern = r"(True Node)\s+(\d+)\s+(\w+)\s+(\d+)\'([\d.]+)"
                match = re.match(pattern, line, re.IGNORECASE)
                if match:
                    planets.append({
                        "planet_name": match.group(1),
                        "positionDegree": match.group(2),
                        "position_sign": zodiac_signs.get(match.group(3), degree_sign_abbr),
                        "position_min": match.group(4),
                        "position_sec": match.group(5),
                    })
                else:
                    planets.append({"error": f"Error parsing line: {line}"})
        else:
            planets.append({"error": "Error parsing output for planets: No lines in the output"})

        # Open the workbook outside of the loop to avoid repeated opening and closing
        try:
            original_path = r'C:\El Camino que Creas\Generador de Informes\Generador de Informes\Generador de Informes.xlsm'
            base, ext = os.path.splitext(original_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # Format: YYYYMMDD_HHMMSS_milliseconds
            copied_file_path = f"{base}_{timestamp}{ext}"
            # wb = xl.Workbooks.Open(file_path)  # Path to your Excel file
            shutil.copyfile(original_path, copied_file_path)
            wb = xl.Workbooks.Open(copied_file_path)  # Path to your Excel file
            
            try:
                
                sheet_name = 'CN y RS (o RL)'  # Replace with your sheet name
                sheet = wb.Sheets(sheet_name)

                # Modify data in the sheet based on the result_data
                for casa, details in result_data.items():
                    if casa.startswith("Casa"):
                        degree_cell = f"E{int(casa[-1]) + 4}"  # Example cell for positionDegree
                        sign_cell = f"D{int(casa[-1]) + 4}"  # Example cell for position_sign
                        min_cell = f"F{int(casa[-1]) + 4}"  # Example cell for position_min
                        sec_cell = f"G{int(casa[-1]) + 4}"  # Example cell for position_sec

                        sheet.Range(degree_cell).Value = details["positionDegree"]
                        sheet.Range(sign_cell).Value = details["position_sign"]
                        sheet.Range(min_cell).Value = details["position_min"]
                        sheet.Range(sec_cell).Value = details["position_sec"]

                # Modify data in the sheet based on the planets
                planet_row_start = 20  # Example starting row for planet data
                for index, planet in enumerate(planets, start=1):
                    if "error" not in planet:
                        
                        # sheet.Range(f"{index + 10}").Value = planet['planet_name']
                        sheet.Range(f"E{index + 10}").Value = planet['positionDegree']
                        sheet.Range(f"D{index + 10}").Value = planet['position_sign']
                        sheet.Range(f"F{index + 10}").Value = planet['position_min']
                        sheet.Range(f"G{index + 10}").Value = planet['position_sec']
                    else:
                        print(planet["error"])
                
                sheet.Range("R26").Value = quiron_parse_output["name"]
                sheet.Range("S26").Value = quiron_parse_output["positionDegree"]
                sheet.Range("T26").Value = quiron_parse_output["position_sign"]
                sheet.Range("U26").Value = quiron_parse_output["position_min"]

                asteroidsList = [quiron_parse_output,lilith_parse_output,result_vertex,cerus_parse_output,pallas_parse_output,juno_parse_output,vesta_parse_output,eris_parse_output,white_moon_output,quaoar_parse_output,sedna_parse_output,varuna_parse_output,nessus_parse_output,waltemath_parse_output,hygeia_parse_output,sylvia_parse_output,hektor_parse_output,europa_parse_output,davida_parse_output,interamnia_parse_output,camilla_parse_output,cybele_parse_output,sol_negro_parse_output,anti_vertex_parse_output,nodo_sur_real_parse_output,sol_negro_real_parse_output,lilith2_parse_output,waltemath_priapus_parse_output,sol_blanco_parse_output,chariklo_parse_output,iris_parse_output,eunomia_parse_output,euphrosyne_output,orcus_parse_output,pholus_parse_output,hermione_parse_output,ixion_parse_output,haumea_parse_output,makemake_parse_output,bamberga_parse_output,patientia_parse_output,thisbe_parse_output,herculina_parse_output,doris_parse_output,ursula_parse_output,eugenia_parse_output,amphitrite_parse_output,diotima_parse_output,fortuna_parse_output,egeria_parse_output,themis_parse_output,aurora_parse_output,alauda_parse_output,aletheia_parse_output,palma_parse_output,nemesis_parse_output,psyche_parse_output,hebe_parse_output,lachesis_parse_output,daphne_parse_output,bertha_parse_output,freia_parse_output,winchester_parse_output,hilda_parse_output,pretoria_parse_output,metis_parse_output,aegle_parse_output,kalliope_parse_output,germania_parse_output,prokne_parse_output,stereoskopia_parse_output,agamemnon_parse_output,alexandra_parse_output,siegena_parse_output,elpis_parse_output,borasisi_parse_output,lempo_parse_output,_1998_26308_parse_output,ceto_parse_output,teharonhiawako_output,_2000_oj67_134860_parse_output,elektra_parse_output,typhon_parse_output,aspasia_parse_output,chicago_parse_output,gyptis_parse_output,diomedes_parse_output,loreley_parse_output,juewa_parse_output,eunike_parse_output,ino_parse_output,ismene_parse_output,merapi_parse_output]
                

                print("Data modified successfully.")
                return jsonify({"message": "Data modified successfully.", "result": result_data, "result2": planets, "asteriods": asteroidsList}), 200
            finally:
                wb.Close(SaveChanges=True)  # Save changes after running macro
        except Exception as e:
            print("Error opening workbook:", e)
            return jsonify({"error": str(e)}), 500
        finally:
            xl.Quit()
   
   
    except Exception as e:
        print("Error initializing Excel:", e)
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

def parse_asteroid_output(asteroid_pholus_output):
    lines = asteroid_pholus_output.splitlines()  # Split by newline characters
    result = {}
    
    
    try:
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 4 or more spaces
            match = re.split(pattern, lines[6])[1]
            name = re.split(pattern, lines[6])[0]
            degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
            degree_match_sign = re.findall(r'[a-zA-Z]+', match)   
            degree_sign = degree_match_sign[0] if degree_match_sign else ""
            degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
            degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
            degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
            degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
            # Only Teharonhiawako Left 
            # When the degree is not found with the first pattern, try the second pattern
            pattern1 = r'\s{2,}'  # Pattern to split by 3 or more spaces
            match1 = re.split(pattern1, lines[6])[1]
            degree_match1 = re.match(r"(\d{1,2})\s\w{2}\s.*", match1)

            result[name] = {
                      "name" : name,
                    "positionDegree": int(degree_match.group(1)) if degree_match else degree_match1.group(1),
                    "position_sign": degree_sign,
                    "position_min": degree_match_min[0],
                    "position_sec": degree_match_min[1] ,                    
                
    
            }
        else:
            result["error"] = "Error parsing output: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result[name]  # Always return a dictionary
