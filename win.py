import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- 1. LÓGICA ESTRATÉGICA (Alineada a Segmentación Novo Nordisk) ---
def obtener_estrategia(nombre, segmento, objetivo, objecion):
    estrategia = {}
    
    # ---------------------------------------------------------
    # SEGMENTO 1: METABÓLICO (Internistas, Alta Adopción GLP-1)
    # ---------------------------------------------------------
    if "Metabólico" in segmento:
        estrategia['wp_1'] = "Posición: Internista. Ve ~25 pacientes/día. Usa: Mounjaro, Wegovy, Ozempic. Aborda la obesidad como enfermedad con >200 comorbilidades."
        estrategia['wp_3'] = "Argumento: Promoción de portafolio WOR. Datos científicos sólidos en todo el síndrome cardiometabólico (inyectable y oral)."
        estrategia['inf_1'] = "Apertura: 'Dr., con su visión integral del paciente y sabiendo que a nivel mundial hay +22 millones de pacientes tratados... hablemos de consolidar el control temprano en pacientes con obesidad y/o ECV.'"
        estrategia['inf_2'] = "Sondeo: 1. ¿En qué momento decide iniciar GLP-1? \n2. ¿Cuándo considera pasar de GLP-1 a otra clase o viceversa?"
        estrategia['inf_3'] = "Beneficio/Valor: Demostrar ≥20% de pérdida de peso en 1 de cada 3 pacientes y ≥57% de reducción de eventos cardiovasculares."
        estrategia['win_1'] = "Victoria: Posiciona el portafolio WOR como base del tratamiento cardiometabólico."

    # ---------------------------------------------------------
    # SEGMENTO 2: ORAL (Cardiólogos, Usa iSGLT2)
    # ---------------------------------------------------------
    elif "Oral" in segmento:
        estrategia['wp_1'] = "Posición: Cardiólogo. Ve ~20 pacientes/día. Usa: iSGLT2. Prioriza glucosa y protección CV. Incomodidad al abordar obesidad frontalmente."
        estrategia['wp_3'] = "Argumento: Transición R -> W. La mejor eficacia entre terapias orales y avanzar hacia semaglutida inyectable para protección CV."
        estrategia['inf_1'] = "Apertura: 'Dr., para sus pacientes con DM2 desde etapas tempranas que prefieren terapias orales, ¿cómo podemos optimizar su protección cardiovascular sin confrontarlos directamente por su peso?'"
        estrategia['inf_2'] = "Sondeo: 1. ¿Qué retos enfrenta para lograr el control glucémico? \n2. ¿Cuál es la reducción promedio de HbA1c que logran hoy?"
        estrategia['inf_3'] = "Beneficio/Valor: Enfatizar 2.2% de reducción de HbA1c y transición suave hacia mayor protección."
        estrategia['win_1'] = "Victoria: Acepta iniciar con Rybelsus y considera la transición a Wegovy por el beneficio CV."

    # ---------------------------------------------------------
    # SEGMENTO 3: PESOCENTRISTA (Bariatras, Urgencia de peso)
    # ---------------------------------------------------------
    else:
        estrategia['wp_1'] = "Posición: Médico Bariatra. Ve ~10 pacientes/día. Usa: Cirugía y Clobenzorex. Busca resultados urgentes en obesidad grado 3."
        estrategia['wp_3'] = "Argumento: Wegovy. Pérdida de peso muy superior en forma oral (DM2) o inyectable, evitando/retrasando la necesidad de cirugía."
        estrategia['inf_1'] = "Apertura: 'Dr., sabiendo que sus pacientes buscan evitar la cirugía por los costos, ¿qué impacto tendría ofrecerles una alternativa médica de alta potencia?'"
        estrategia['inf_2'] = "Sondeo: 1. ¿Qué obstáculos enfrenta para que sus pacientes completen el tratamiento? \n2. ¿Cuáles son sus metas a corto y mediano plazo?"
        estrategia['inf_3'] = "Beneficio/Valor: Enfatizar rapidez y el dato duro de ≥20% de pérdida de peso en 1 de cada 3 pacientes."
        estrategia['win_1'] = "Victoria: Desplaza a los anoréxicos tradicionales (Clobenzorex) y posiciona Wegovy como primera línea pre-quirúrgica."

    # --- ELEMENTOS COMUNES ---
    estrategia['wp_2'] = f"Objeción esperada: '{objecion}'. \nAplicar Metodología PACT (Pausa, Agradece, Clarifica, Transforma)."
    estrategia['wp_4'] = f"Meta: {objetivo}."
    estrategia['nba_1'] = "Plan de Acción: Obtener compromiso de inicio en al menos 2 pacientes con el perfil conversado."
    estrategia['nba_2'] = "Plan B: Enviar evidencia clínica (ej. reducción de eventos ≥57%) y agendar revisión de caso clínico."
    
    return estrategia