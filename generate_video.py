"""
generate_video.py — Pipeline de generación de video de capacitación QA TeleVentas Despegar
Requiere: pip install -r requirements.txt
          apt-get install -y poppler-utils ffmpeg

Uso:
    python generate_video.py --script narration_script.txt --output capacitacion_qa_tlv.mp4

Fases:
    1. Extrae páginas PDF como imágenes (300 DPI)
    2. Genera slides de glosario con Pillow
    3. Genera audio TTS por slide (edge-tts con fallback a gTTS)
    4. Ensambla video final con moviepy (1920x1080, H.264, crossfade 0.4s)
"""

import os
import sys
import re
import asyncio
import argparse
import textwrap
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
import pymupdf

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN GLOBAL
# ──────────────────────────────────────────────────────────────────────────────

WIDTH, HEIGHT = 1920, 1080
DESPEGAR_PURPLE = (91, 46, 250)      # #5B2EFA
DESPEGAR_DARK   = (30, 15, 80)       # dark variant for backgrounds
WHITE           = (255, 255, 255)
DARK_GRAY       = (50, 50, 50)
LIGHT_GRAY      = (240, 240, 245)

CROSSFADE_DURATION = 0.4             # seconds
SLIDE_PAUSE        = 0.8             # silence after narration ends
CHAPTER_CARD_DURATION = 3.0          # display time for chapter title cards
TTS_LANGUAGE       = "es-MX"
TTS_VOICE_EDGE     = "es-MX-DaliaNeural"   # female, Mexican Spanish

PDF_DPI = 150                        # lower DPI to reduce image sizes (still HD-compatible)

# Paths
BASE_DIR   = Path(__file__).parent
AUDIO_DIR  = BASE_DIR / "audio_slides"
IMAGE_DIR  = BASE_DIR / "slide_images"
OUTPUT_MP4 = BASE_DIR / "capacitacion_qa_tlv.mp4"

PDF_QA     = BASE_DIR / "Capacitación de QA.pdf"
PDF_CXC    = BASE_DIR / "Política de cobro por errores en ventas - CXC (1).pdf"
XLSX_GLOS  = BASE_DIR / "Glosario TLV  _ OCT (1).xlsx"

# Fonts (system fonts, fallback to default PIL)
FONT_PATHS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    "/usr/share/fonts/truetype/ubuntu/Ubuntu-B.ttf",
]
FONT_REGULAR_PATHS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    "/usr/share/fonts/truetype/ubuntu/Ubuntu-R.ttf",
]


def get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    paths = FONT_PATHS if bold else FONT_REGULAR_PATHS
    for p in paths:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


# ──────────────────────────────────────────────────────────────────────────────
# SLIDE DEFINITIONS
# Each entry: (slide_id, source, narration)
# source: ("pdf_qa", page_index) | ("pdf_cxc", page_index) | ("glossary", terms_list)
#         | ("chapter_card", title_text) | ("intro",) | ("outro",)
# ──────────────────────────────────────────────────────────────────────────────

SLIDES = [
    # ── INTRO ──────────────────────────────────────────────────────────────
    {
        "id": "000-intro",
        "source": ("intro",),
        "narration": (
            "Bienvenido al programa de capacitación de TeleVentas de Despegar. "
            "Este video fue diseñado especialmente para ti, como nuevo agente de ventas, "
            "para que llegues a tu sesión de onboarding con contexto sólido sobre tres temas clave: "
            "el proceso de evaluación de QA, la política de cobro por errores CXC, "
            "y los términos del glosario de TeleVentas. Vamos a empezar."
        ),
    },

    # ── CAPÍTULO 1 ──────────────────────────────────────────────────────────
    {
        "id": "ch1-title",
        "source": ("chapter_card", "Capítulo 1\nCapacitación QA\nMejora Continua"),
        "narration": (
            "Ahora vamos a hablar del Capítulo 1: el proceso de evaluación de calidad de TeleVentas. "
            "Vas a conocer cómo está estructurada la rúbrica, qué se evalúa en cada tipo de llamada "
            "y cuáles son las actividades del equipo de QA."
        ),
    },
    {
        "id": "1-1",
        "source": ("pdf_qa", 0),   # page 1
        "narration": (
            "El área de Mejora Continua QA de TeleVentas es responsable de evaluar la calidad "
            "de cada llamada o chat que atiendes. Su objetivo es garantizar que cada cliente "
            "reciba atención de excelencia y que el proceso de venta cumpla los estándares de Despegar. "
            "Como agente nuevo, entender cómo funciona esta evaluación desde el inicio "
            "te dará una ventaja importante."
        ),
    },
    {
        "id": "1-2",
        "source": ("pdf_qa", 1),   # page 2: NO VENTA IN/Virtual
        "narration": (
            "En las llamadas sin venta, cuando el cliente no concreta una compra, el agente "
            "igual es evaluado. Para llamadas IN o Virtual, la rúbrica tiene seis ejes: "
            "Atención y servicio, Discovery, Técnicas de venta, Estrategia comercial, Productos "
            "cotizados, y Faltas Graves. Presta especial atención a las Faltas Graves y a todos "
            "los ítems marcados en rojo: son Errores Críticos y llevan tu evaluación directamente "
            "a cero, sin importar qué tan bien hayas hecho el resto de la llamada."
        ),
    },
    {
        "id": "1-3",
        "source": ("pdf_qa", 2),   # page 3: VENTA IN/Virtual
        "narration": (
            "Cuando sí se cierra una venta, los criterios cambian. Se agrega la sección de "
            "Confirmación, que es el momento más crítico del proceso: deletreo y re-confirmación "
            "de datos de pasajeros, revisión de cargos, confirmación de productos reservados y "
            "ofrecimiento de facturación. Todos estos ítems están marcados en rojo, es decir, "
            "son Errores Críticos. Si fallas en confirmar correctamente los datos del pasajero "
            "o los cargos, tu evaluación cae a cero. También debes documentar la gestión en Fénix."
        ),
    },
    {
        "id": "1-4",
        "source": ("pdf_qa", 3),   # page 4: NO VENTA WhatsApp
        "narration": (
            "El canal de WhatsApp tiene su propia rúbrica. En interacciones sin venta, se "
            "evalúan el saludo y despedida, el manejo del chat, la tipificación correcta, "
            "los datos básicos del cliente y la ortografía y redacción. En WhatsApp no hay voz: "
            "la forma en que escribes representa la imagen de Despegar. Los tiempos de respuesta "
            "máximos son críticos: para agentes fijos en WhatsApp el límite es de 15 minutos, "
            "y para agentes multiskill, 45 minutos."
        ),
    },
    {
        "id": "1-5",
        "source": ("pdf_qa", 4),   # page 5: VENTA WhatsApp
        "narration": (
            "En ventas por WhatsApp, los criterios de Faltas Graves incluyen elementos que "
            "en llamadas IN no aparecen como críticos: el cross selling, los productos cotizados, "
            "las alternativas ofrecidas, el cierre de venta y la facturación. Todos son Errores Críticos. "
            "La confirmación en WhatsApp se realiza enviando un PDF completo al cliente con los datos "
            "del servicio, y el cliente debe confirmar por escrito que todo es correcto "
            "antes de generar la reserva."
        ),
    },
    {
        "id": "1-6",
        "source": ("pdf_qa", 6),   # page 7: Actividades QA
        "narration": (
            "¿Qué hace el equipo de QA en su día a día? Cuatro actividades principales. "
            "Primero, evaluaciones con inteligencia artificial, que monitorean un porcentaje "
            "de llamadas de manera automática. Segundo, auditorías de llamadas PI o legales, "
            "cuando hay reclamaciones de clientes. Tercero, auditorías internas operativas "
            "para verificar la calidad del proceso comercial. Y cuarto, revisión de pagos en destino. "
            "Como agente, puedes ser auditado en cualquier momento y en cualquier llamada."
        ),
    },
    {
        "id": "1-7",
        "source": ("pdf_qa", 7),   # page 8: Proceso de Revisión
        "narration": (
            "Si recibes una evaluación con la que no estás de acuerdo, existe un proceso formal "
            "de revisión en cuatro pasos. Uno: tienes 48 horas desde que recibes la evaluación "
            "para enviar tu refute, es decir, tu inconformidad. Dos: tu supervisor escucha la llamada. "
            "Tres: el analista de QA recibe el refute y tiene hasta 72 horas para revisarlo. "
            "Cuatro: el analista te responde por correo confirmando si el refute fue aceptado o rechazado. "
            "Ese plazo de 48 horas para refutar no se extiende."
        ),
    },

    # ── CAPÍTULO 2 ──────────────────────────────────────────────────────────
    {
        "id": "ch2-title",
        "source": ("chapter_card", "Capítulo 2\nPolítica CXC\nCobro por Errores en Ventas"),
        "narration": (
            "Ahora vamos al Capítulo 2: la Política CXC. Este tema es importante conocerlo "
            "no para intimidarse, sino para entender las consecuencias de errores graves en ventas "
            "y cómo funciona el proceso de manera transparente."
        ),
    },
    {
        "id": "2-1",
        "source": ("pdf_cxc", 3),   # page 4: Objetivo
        "narration": (
            "El objetivo de la Política CXC es definir un proceso claro para aplicar descuentos "
            "salariales cuando un agente causa una pérdida económica comprobada. No cualquier "
            "error lleva a un descuento. Para que aplique, deben cumplirse tres condiciones: "
            "que haya negligencia, fraude o incumplimiento de políticas conocidas; "
            "que exista evidencia documentada como audios o registros en el CRM; "
            "y que la pérdida sea directa, cuantificable y atribuible al agente. "
            "Ejemplos: ventas fraudulentas, manipulación de sistemas o uso indebido de descuentos."
        ),
    },
    {
        "id": "2-2",
        "source": ("pdf_cxc", 4),   # page 5: Proceso General
        "narration": (
            "¿Cómo inicia todo? Con una reclamación del cliente. El equipo de posventa crea "
            "un caso en Fénix asociado a una PI, un Pedido de Información. Con esa PI, calidad "
            "audita toda la gestión, desde el primer contacto hasta el cierre. Si el error es "
            "atribuible al agente, se notifica formalmente al agente, al supervisor y al manager. "
            "Después, Revenue Management confirma el monto de la pérdida con el proveedor "
            "y lo reporta a mes vencido. El proceso puede extenderse varios meses."
        ),
    },
    {
        "id": "2-3",
        "source": ("pdf_cxc", 5),   # page 6: Proceso de Descuento
        "narration": (
            "Si el Comité de Incidencias confirma la responsabilidad del agente, se autoriza "
            "el descuento. Este se aplica conforme a la Ley Federal del Trabajo. El equipo de "
            "Nómina, junto con el agente, define si el monto se descuenta en una sola exhibición "
            "o en cuotas. El agente firma un documento de aceptación; si hay desacuerdo, puede "
            "iniciar un proceso de conciliación. Punto clave: si hay dos pérdidas pendientes, "
            "no se pueden aplicar al mismo tiempo. Primero se cierra una y luego se firma la segunda."
        ),
    },
    {
        "id": "2-4",
        "source": ("pdf_cxc", 6),   # page 7: Tabla de Descuentos
        "narration": (
            "Esta es la tabla de descuentos. Con 1 a 3 meses de antigüedad, el tope máximo a "
            "cubrir es de 500 dólares. En tu primera PI: cero descuento. Segunda PI: 30 por ciento. "
            "Tercera PI: 50 por ciento. Con más de 3 meses de antigüedad, el tope sube a mil dólares, "
            "y los porcentajes son 50, 75 y 100 por ciento para primera, segunda y tercera PI. "
            "Lo que exceda esos topes lo cubre Grupo Despegar. Esta política aplica para México "
            "desde el 1 de julio de 2025."
        ),
    },
    {
        "id": "2-5",
        "source": ("pdf_cxc", 8),   # page 9: Comité de Incidencias
        "narration": (
            "¿Quiénes toman estas decisiones? El Comité de Incidencias está formado por "
            "el Senior Telesales Manager, el B2B2C Manager, la Operations Manager de Training "
            "y Quality, la Telesales Operations Manager, el HRBP Specialist de Recursos Humanos "
            "y el Payroll Specialist de Nómina. Este comité existe para garantizar que ninguna "
            "decisión se tome de manera arbitraria. Todo el proceso queda documentado "
            "y firmado por los involucrados."
        ),
    },
    {
        "id": "2-6",
        "source": ("pdf_cxc", 11),  # page 12: Proceso Completo
        "narration": (
            "El flujo completo es este: el cliente genera la reclamación, posventa crea la PI, "
            "calidad audita la venta, Revenue consolida y reporta la pérdida a Operaciones "
            "y Recursos Humanos. El Comité se reúne para confirmar el modelo de descuento; "
            "el supervisor y el HRBP firman la carta CXC con el agente; "
            "y finalmente Nómina aplica el descuento. Es un proceso ordenado, con múltiples "
            "revisiones antes de cualquier impacto económico."
        ),
    },

    # ── CAPÍTULO 3: GLOSARIO ────────────────────────────────────────────────
    {
        "id": "ch3-title",
        "source": ("chapter_card", "Capítulo 3\nGlosario\nTeleVentas"),
        "narration": (
            "Llegamos al Capítulo 3: el Glosario de TeleVentas. Vamos a repasar los términos "
            "y conceptos clave que escucharás todos los días. "
            "Conocerlos desde ahora te dará una ventaja enorme desde tu primer día."
        ),
    },
    {
        "id": "3-1",
        "source": ("glossary", [
            ("ENC", "Encuadre Normal de Calidad",
             "Errores que reducen el puntaje de forma parcial, según el peso del ítem en la rúbrica."),
            ("EC", "Error Crítico",
             "Ítems marcados en ROJO en la rúbrica. Llevan la evaluación a CERO de inmediato, sin importar el resto."),
        ]),
        "narration": (
            "Hay dos tipos de error en la evaluación de QA. ENC, Encuadre Normal de Calidad, "
            "son errores que afectan tu puntaje de manera parcial, según el peso de ese ítem. "
            "EC, Error Crítico, son los ítems marcados en rojo en la rúbrica. "
            "Si cometes un EC, tu evaluación cae a cero de inmediato, "
            "sin importar qué tan bien hayas hecho el resto de la llamada. "
            "Conocer cuáles ítems son EC es lo primero que debes tener muy claro."
        ),
    },
    {
        "id": "3-2",
        "source": ("glossary", [
            ("Discovery", "",
             "Fase de la llamada donde identificas las necesidades del cliente antes de cotizar."),
            ("Data Básica", "",
             "Información mínima para cotizar: fechas, origen, destino, número de pasajeros y tipo de servicio."),
            ("Sondeo", "",
             "Indagación sutil de motivaciones, preferencias y presupuesto del cliente para ofrecer lo que realmente necesita."),
        ]),
        "narration": (
            "Discovery es la fase de la llamada donde identificas las necesidades del cliente. "
            "Dentro del Discovery hay dos elementos clave. Data básica: la información mínima "
            "para cotizar, que incluye fechas, origen, destino, número de pasajeros y tipo de servicio; "
            "una vez obtenida, debes re-confirmarla con el cliente. Sondeo: va más allá de los datos, "
            "es indagar de manera sutil las motivaciones del viaje, las preferencias y el presupuesto "
            "del cliente, para poder ofrecerle lo que realmente necesita."
        ),
    },
    {
        "id": "3-3",
        "source": ("glossary", [
            ("Manejo de Llamada", "",
             "Estructura general: saludo, despedida, escucha activa, lenguaje claro y tono de voz adecuado."),
            ("Tiempos de Espera", "",
             "Máximo 1 minuto sin interactuar en llamadas. En WhatsApp: 15 min (agentes fijos) / 45 min (multiskill)."),
        ]),
        "narration": (
            "El Manejo de llamada evalúa tu estructura general: el saludo, la despedida, "
            "la escucha activa, el lenguaje claro y el tono de voz. También incluye los tiempos "
            "de espera. El máximo de silencio permitido sin interactuar con el cliente es de 1 minuto. "
            "Si necesitas consultar algo, avísale al cliente y, al regresar, agradécele el tiempo. "
            "En WhatsApp: 15 minutos para agentes fijos y 45 para agentes multiskill. "
            "Superar estos tiempos puede ser un Error Crítico."
        ),
    },
    {
        "id": "3-4",
        "source": ("glossary", [
            ("Cross Selling", "Venta Cruzada",
             "Ofrecer servicios adicionales para complementar el viaje: traslados, seguros, tours. Se ofrece activamente o incluyendo en la cotización."),
            ("Upselling", "",
             "Invitar al cliente a subir de categoría: habitación estándar → suite; traslado compartido → privado."),
        ]),
        "narration": (
            "Cross selling o venta cruzada es ofrecer servicios adicionales al que el cliente "
            "ya solicitó, para complementar su experiencia. Por ejemplo: el cliente quiere vuelo "
            "y tú le ofreces traslado o seguro de asistencia. Se puede hacer ofreciendo activamente "
            "con beneficios, o incluyendo el servicio en el total de la cotización. Upselling es "
            "invitar al cliente a subir de categoría, por ejemplo, de habitación estándar a suite "
            "vista al mar. Ambas técnicas generan valor al cliente y aumentan el ticket de la reserva."
        ),
    },
    {
        "id": "3-5",
        "source": ("glossary", [
            ("Dream Phase", "",
             "Explicar con claridad y entusiasmo mínimo 3 beneficios o atributos del hotel o destino para despertar el deseo de compra."),
            ("Direccionamiento", "",
             "Orientar al cliente hacia productos prioritarios del negocio: circuitos, cabinas, promociones especiales."),
        ]),
        "narration": (
            "Dream Phase es la capacidad de despertar el deseo de compra explicando con claridad "
            "los beneficios y atributos del producto, especialmente del hotel o destino. "
            "Se evalúa que menciones al menos 3 características relevantes. No es leer una ficha técnica, "
            "sino hablar con conocimiento y entusiasmo. Direccionamiento aplica cuando el negocio "
            "tiene una prioridad comercial específica: circuitos, cabinas, promociones. Tu responsabilidad "
            "es conocer esas prioridades y dirigir al cliente hacia ellas de manera natural."
        ),
    },
    {
        "id": "3-6",
        "source": ("glossary", [
            ("Confirmación", "Speech OBLIGATORIO",
             "Antes de reservar, debes leer el speech obligatorio y obtener la aceptación del cliente. EC si se omite."),
            ("Deletreo y Re-confirmación", "⚠ EC",
             "Confirmar uno a uno: nombres y apellidos (fonéticamente), fecha de nacimiento e identificación oficial."),
        ]),
        "narration": (
            "La Confirmación es el momento más crítico de una llamada de venta, y es un Error Crítico "
            "si se omite o se hace mal. Antes de reservar, debes leer el speech obligatorio: "
            "Señor o Señora, pasaremos a la parte de la confirmación del servicio a reservar. "
            "Le pido su completa atención para evitar cualquier error. Una vez aceptada su compra, "
            "queda sujeto a los términos y condiciones de la página. Después confirmas uno por uno: "
            "nombres y apellidos deletreando fonéticamente, fecha de nacimiento e identificación oficial. "
            "Un error aquí lleva la evaluación a cero."
        ),
    },
    {
        "id": "3-7",
        "source": ("glossary", [
            ("Cargos", "⚠ EC",
             "Mencionar método de pago, cargos adicionales (tasas, IVA, resort fee). Hablar siempre con aproximados."),
            ("Productos Reservados", "⚠ EC",
             "Informar completo: fechas, detalles del servicio, pasajeros, políticas de cancelación y precio."),
            ("Facturación", "⚠ EC",
             "Ofrecer factura desde el momento de la reservación tomando y confirmando los datos correctamente."),
        ]),
        "narration": (
            "Dentro de la Confirmación también revisas tres cosas más, todas de tipo Error Crítico. "
            "Cargos: menciona el método de pago y los cargos adicionales como tasas, impuestos, "
            "resort fee o IVA; habla siempre con aproximados, no garantices tarifas exactas. "
            "Productos reservados: proporciona información completa del servicio, incluyendo fechas, "
            "detalles del destino, número de pasajeros, políticas de cancelación y precio. "
            "Facturación: ofrece generar factura desde el momento de la reservación "
            "tomando los datos correctamente."
        ),
    },
    {
        "id": "3-8",
        "source": ("glossary", [
            ("Faltas Graves", "Temas comportamentales — ⚠ EC",
             "• Temas actitudinales: sarcasmo, ironía, pasivo-agresivo\n"
             "• Falta de respeto: palabras despectivas, bostezar, comer en llamada\n"
             "• Evasión del servicio: negarse a ayudar\n"
             "• Malas prácticas: cobrar extra sin informar, bajar reservas sin autorización\n"
             "• Agente virtual: cámara con entorno inadecuado, sin uniforme"),
        ]),
        "narration": (
            "Las Faltas Graves son siempre Error Crítico, sin excepción. "
            "Temas actitudinales: actitud pasivo-agresiva, sarcasmo o burla disimulada. "
            "Falta de respeto: palabras despectivas, bostezar o comer durante la llamada. "
            "Evasión del servicio: negarse a ayudar o no tener disposición. "
            "Malas prácticas: cobrar un servicio extra sin informarle al cliente, "
            "bajar una reserva sin su autorización, o bajar reservas desde la plataforma online. "
            "Agente virtual: cámara encendida con entorno inadecuado, mala postura o sin uniforme."
        ),
    },
    {
        "id": "3-9",
        "source": ("glossary", [
            ("Pérdida del Llamado", "⚠ EC",
             "PROHIBIDO cortar la llamada de manera deliberada y sin justificación.\nSi es necesario: 'Por cuestión de calidad en el servicio, procedo a liberar la llamada.'"),
            ("Propone Agendamiento", "Call Back — ⚠ EC",
             "Ofrecer seguimiento por call back proponiendo fecha y hora concretas. El seguimiento solo por correo NO es válido."),
        ]),
        "narration": (
            "Pérdida del llamado significa cortar la llamada de manera deliberada y sin justificación. "
            "Está completamente prohibido. Si por una razón técnica necesitas liberar la llamada, "
            "debes anunciarlo: Por cuestión de calidad en el servicio, procedo a liberar la llamada. "
            "Y debes regresar el llamado al cliente. Propone agendamiento: cuando el cliente no está "
            "listo para comprar, ofrece un seguimiento por call back proponiendo fecha y hora concretas. "
            "El seguimiento solo por correo no es válido."
        ),
    },
    {
        "id": "3-10",
        "source": ("glossary", [
            ("Tipificación", "",
             "Registrar en el sistema el motivo real de la llamada o chat de manera correcta y completa al finalizar cada gestión."),
            ("Fénix", "CRM Principal",
             "CRM donde documentas la gestión: URL de llamada (pago en destino), correo con nombres confirmados (+5 pasajeros)."),
            ("Sócrates", "CRM Complementario",
             "Sistema donde registras la tipificación. Consultado por QA al auditar tus llamadas."),
        ]),
        "narration": (
            "Al finalizar cada gestión debes tipificar, es decir, registrar en el sistema "
            "el motivo real de la llamada o chat de manera correcta y completa. Fénix es el CRM "
            "principal de Despegar donde documentas la gestión: pegas el URL de la llamada en casos "
            "de pago en destino, o el correo con nombres confirmados cuando hay más de 5 pasajeros. "
            "Sócrates es el sistema complementario donde también registras la tipificación. "
            "El equipo de QA consulta ambos sistemas al auditar una llamada."
        ),
    },
    {
        "id": "3-11",
        "source": ("glossary", [
            ("PI", "Pedido de Información",
             "Ticket generado cuando un cliente hace una reclamación. Dispara el proceso de auditoría de QA."),
            ("MSI", "Meses Sin Intereses",
             "Opción de pago a ofrecer activamente. Ejemplo: 3, 6, 9 o 12 MSI según el banco disponible."),
            ("VF", "Virtual Fuerza",
             "Modalidad de trabajo remoto / virtual de los agentes de TeleVentas."),
            ("TLV", "TeleVentas",
             "El área donde trabajas. Sinónimo de TeleVentas en la operación de Despegar."),
        ]),
        "narration": (
            "Cuatro abreviaturas que escucharás con mucha frecuencia. PI, Pedido de Información: "
            "es el ticket que se genera cuando un cliente reclama y dispara el proceso de auditoría. "
            "MSI, Meses Sin Intereses: opción de pago que debes ofrecer activamente, mencionando "
            "los bancos disponibles y las opciones de 3, 6, 9 o 12 meses. "
            "VF es Virtual Fuerza, la modalidad de trabajo remoto de los agentes. "
            "Y TLV simplemente significa TeleVentas, el área donde trabajas."
        ),
    },
    {
        "id": "3-12",
        "source": ("glossary", [
            ("Manejo de Objeciones", "⚠ EC en modalidades de venta",
             "Investigar la verdadera objeción del cliente y responder asertivamente para orientarlo al cierre. Mínimo 1 vez."),
            ("Ofrece Alternativas", "⚠ EC en NO VENTA",
             "Si no hay disponibilidad o el cliente rechaza la primera opción, busca alternativas en fechas, destinos, tarifas o categorías."),
            ("Propone Cierre de Venta", "⚠ EC en WhatsApp",
             "Generar activamente la propuesta de cierre usando herramientas como cupones o promociones disponibles."),
        ]),
        "narration": (
            "Para cerrar el glosario, tres técnicas de cierre. Manejo de objeciones: cuando el "
            "cliente duda, investiga la verdadera razón y responde asertivamente para orientarlo "
            "a la decisión de compra, mínimo una vez. Ofrece alternativas: si no hay disponibilidad "
            "o el cliente rechaza la primera opción, busca opciones en fechas, destinos, tarifas "
            "o categorías. Propone cierre de venta: especialmente en WhatsApp, genera activamente "
            "la propuesta de cierre usando cupones o promociones disponibles. Estas tres acciones "
            "marcan la diferencia entre un agente promedio y uno de alto desempeño."
        ),
    },

    # ── OUTRO ───────────────────────────────────────────────────────────────
    {
        "id": "outro",
        "source": ("outro",),
        "narration": (
            "Y con esto concluye tu capacitación previa al onboarding. "
            "Recuerda los tres pilares que revisamos hoy: la rúbrica de evaluación de QA, "
            "la política de cobro por errores CXC, y el glosario de términos clave. "
            "Ahora llegas a tu sesión de bienvenida con contexto. "
            "Esa sesión es tu espacio para resolver dudas y practicar con casos reales. "
            "¡Mucho éxito en esta nueva etapa! Estamos muy contentos de tenerte en el equipo de Despegar."
        ),
    },
]


# ──────────────────────────────────────────────────────────────────────────────
# SLIDE IMAGE GENERATORS
# ──────────────────────────────────────────────────────────────────────────────

def make_intro_slide() -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), DESPEGAR_DARK)
    draw = ImageDraw.Draw(img)
    # Top purple bar
    draw.rectangle([(0, 0), (WIDTH, 140)], fill=DESPEGAR_PURPLE)
    font_small = get_font(32)
    draw.text((60, 52), "Despegar TeleVentas", font=font_small, fill=WHITE)
    # Main title
    font_big = get_font(96, bold=True)
    title = "Capacitación de Ingreso"
    bb = draw.textbbox((0, 0), title, font=font_big)
    draw.text(((WIDTH - (bb[2]-bb[0]))//2, 300), title, font=font_big, fill=WHITE)
    font_sub = get_font(52)
    sub = "QA · Política CXC · Glosario TLV"
    bb2 = draw.textbbox((0, 0), sub, font=font_sub)
    draw.text(((WIDTH - (bb2[2]-bb2[0]))//2, 450), sub, font=font_sub, fill=(200, 190, 255))
    # Bottom bar
    draw.rectangle([(0, HEIGHT-80), (WIDTH, HEIGHT)], fill=DESPEGAR_PURPLE)
    font_xs = get_font(26)
    draw.text((60, HEIGHT-58), "Mejora Continua QA · TeleVentas", font=font_xs, fill=WHITE)
    return img


def make_outro_slide() -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), DESPEGAR_DARK)
    draw = ImageDraw.Draw(img)
    draw.rectangle([(0, 0), (WIDTH, 140)], fill=DESPEGAR_PURPLE)
    font_small = get_font(32)
    draw.text((60, 52), "Despegar TeleVentas", font=font_small, fill=WHITE)
    font_big = get_font(110, bold=True)
    text = "¡Bienvenido al equipo!"
    bb = draw.textbbox((0, 0), text, font=font_big)
    draw.text(((WIDTH - (bb[2]-bb[0]))//2, 340), text, font=font_big, fill=WHITE)
    font_sub = get_font(48)
    sub = "#SomosParteD!"
    bb2 = draw.textbbox((0, 0), sub, font=font_sub)
    draw.text(((WIDTH - (bb2[2]-bb2[0]))//2, 500), sub, font=font_sub, fill=(200, 190, 255))
    draw.rectangle([(0, HEIGHT-80), (WIDTH, HEIGHT)], fill=DESPEGAR_PURPLE)
    return img


def make_chapter_card(title_text: str) -> Image.Image:
    img = Image.new("RGB", (WIDTH, HEIGHT), DESPEGAR_PURPLE)
    draw = ImageDraw.Draw(img)
    # Decorative diagonal stripe
    for i in range(0, 500, 60):
        draw.polygon(
            [(WIDTH-i, 0), (WIDTH-i+40, 0), (WIDTH, i+40), (WIDTH, i)],
            fill=(255, 255, 255, 20)
        )
    lines = title_text.split("\n")
    fonts = [get_font(52, bold=False), get_font(88, bold=True), get_font(60, bold=False)]
    colors = [(220, 210, 255), WHITE, (220, 210, 255)]
    total_height = sum(
        draw.textbbox((0, 0), l, font=fonts[min(i, len(fonts)-1)])[3]
        for i, l in enumerate(lines)
    ) + 30 * (len(lines) - 1)
    y = (HEIGHT - total_height) // 2 - 40
    for i, line in enumerate(lines):
        f = fonts[min(i, len(fonts)-1)]
        c = colors[min(i, len(colors)-1)]
        bb = draw.textbbox((0, 0), line, font=f)
        x = (WIDTH - (bb[2]-bb[0])) // 2
        draw.text((x, y), line, font=f, fill=c)
        y += (bb[3]-bb[1]) + 30
    return img


def make_glossary_slide(terms: list) -> Image.Image:
    """Generate a glossary slide with 1–4 terms."""
    img = Image.new("RGB", (WIDTH, HEIGHT), WHITE)
    draw = ImageDraw.Draw(img)

    # Purple header bar
    BAR_H = 90
    draw.rectangle([(0, 0), (WIDTH, BAR_H)], fill=DESPEGAR_PURPLE)
    font_bar = get_font(30)
    draw.text((50, 28), "Glosario TeleVentas  ·  Capacitación de Ingreso", font=font_bar, fill=WHITE)

    # Content area
    content_top = BAR_H + 30
    content_bottom = HEIGHT - 30
    slot_height = (content_bottom - content_top) // len(terms)

    for idx, item in enumerate(terms):
        term_name, term_subtitle, definition = item
        y_start = content_top + idx * slot_height
        y_end = y_start + slot_height - 20

        # Light background for alternating slots
        if idx % 2 == 1:
            draw.rectangle([(20, y_start), (WIDTH-20, y_end)], fill=LIGHT_GRAY)

        # Term name
        font_term = get_font(min(54, max(36, 54 - len(term_name))), bold=True)
        term_display = term_name
        if term_subtitle:
            term_display = f"{term_name}  ·  {term_subtitle}"
        draw.text((50, y_start + 18), term_display, font=font_term, fill=DESPEGAR_PURPLE)

        # Definition (wrapped)
        font_def = get_font(28)
        max_chars = 90
        def_lines = []
        for dline in definition.split("\n"):
            wrapped = textwrap.wrap(dline, width=max_chars)
            def_lines.extend(wrapped if wrapped else [""])
        max_def_lines = max(1, (y_end - y_start - 80) // 36)
        def_lines = def_lines[:max_def_lines]
        y_def = y_start + 76
        for dl in def_lines:
            draw.text((70, y_def), dl, font=font_def, fill=DARK_GRAY)
            y_def += 36

        # Separator line
        if idx < len(terms) - 1:
            draw.line([(30, y_end + 8), (WIDTH-30, y_end + 8)], fill=(200, 200, 220), width=2)

    # Bottom purple accent
    draw.rectangle([(0, HEIGHT-6), (WIDTH, HEIGHT)], fill=DESPEGAR_PURPLE)
    return img


def pdf_page_to_image(pdf_path: Path, page_index: int) -> Image.Image:
    doc = pymupdf.open(str(pdf_path))
    page = doc[page_index]
    mat = pymupdf.Matrix(PDF_DPI / 72, PDF_DPI / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    # Fit to 1920x1080 with letterboxing
    img.thumbnail((WIDTH, HEIGHT), Image.LANCZOS)
    bg = Image.new("RGB", (WIDTH, HEIGHT), (20, 10, 50))
    x = (WIDTH - img.width) // 2
    y = (HEIGHT - img.height) // 2
    bg.paste(img, (x, y))
    return bg


# ──────────────────────────────────────────────────────────────────────────────
# TTS AUDIO GENERATION
# ──────────────────────────────────────────────────────────────────────────────

async def generate_audio_edge(text: str, out_path: Path) -> bool:
    try:
        import edge_tts
        communicate = edge_tts.Communicate(text, voice=TTS_VOICE_EDGE)
        await communicate.save(str(out_path))
        return True
    except Exception as e:
        print(f"  [edge-tts] failed: {e}")
        return False


def generate_audio_gtts(text: str, out_path: Path) -> bool:
    try:
        from gtts import gTTS
        tts = gTTS(text=text, lang="es", tld="com.mx", slow=False)
        tts.save(str(out_path))
        return True
    except Exception as e:
        print(f"  [gTTS] failed: {e}")
        return False


async def generate_audio(text: str, out_path: Path):
    if out_path.exists():
        print(f"  [audio] reusing {out_path.name}")
        return
    print(f"  [audio] generating {out_path.name} ...")
    ok = await generate_audio_edge(text, out_path)
    if not ok:
        ok = generate_audio_gtts(text, out_path)
    if not ok:
        raise RuntimeError(f"TTS failed for slide — both edge-tts and gTTS failed.")


# ──────────────────────────────────────────────────────────────────────────────
# VIDEO ASSEMBLY WITH MOVIEPY
# ──────────────────────────────────────────────────────────────────────────────

def get_audio_duration(audio_path: Path) -> float:
    from moviepy.editor import AudioFileClip
    clip = AudioFileClip(str(audio_path))
    dur = clip.duration
    clip.close()
    return dur


def build_video_clip(image: Image.Image, audio_path: Path, is_chapter_card: bool = False):
    """Return a moviepy VideoClip for a single slide."""
    from moviepy.editor import ImageClip, AudioFileClip, CompositeVideoClip
    from moviepy.editor import ColorClip

    img_array = np.array(image.resize((WIDTH, HEIGHT), Image.LANCZOS))
    audio = AudioFileClip(str(audio_path))
    narration_dur = audio.duration

    if is_chapter_card:
        display_dur = max(CHAPTER_CARD_DURATION, narration_dur)
    else:
        display_dur = narration_dur + SLIDE_PAUSE

    video = ImageClip(img_array, duration=display_dur)
    video = video.set_audio(audio)
    return video


def assemble_video(clips: list, output_path: Path):
    """Concatenate clips with crossfade transitions and save as MP4."""
    from moviepy.editor import concatenate_videoclips

    print(f"\n[video] Assembling {len(clips)} clips ...")
    # Apply crossfade
    faded = []
    for i, clip in enumerate(clips):
        if i == 0:
            # Fade in at start
            clip = clip.fadein(0.8)
        if i == len(clips) - 1:
            # Fade out at end
            clip = clip.fadeout(0.8)
        faded.append(clip)

    final = concatenate_videoclips(faded, method="compose", padding=-CROSSFADE_DURATION)
    print(f"[video] Writing {output_path} ...")
    final.write_videofile(
        str(output_path),
        fps=24,
        codec="libx264",
        audio_codec="aac",
        bitrate="4000k",
        audio_bitrate="128k",
        preset="medium",
        threads=4,
        logger="bar",
    )
    print(f"\n[video] Done → {output_path}")


# ──────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ──────────────────────────────────────────────────────────────────────────────

async def main(output_path: Path = OUTPUT_MP4):
    AUDIO_DIR.mkdir(exist_ok=True)
    IMAGE_DIR.mkdir(exist_ok=True)

    clips = []
    total_dur = 0.0

    for slide in SLIDES:
        sid = slide["id"]
        source = slide["source"]
        narration = slide["narration"]
        is_chapter_card = source[0] == "chapter_card"

        print(f"\n[slide] {sid}")

        # ── 1. Generate or load image ──────────────────────────────────────
        img_path = IMAGE_DIR / f"{sid}.png"
        if img_path.exists():
            print(f"  [image] reusing {img_path.name}")
            img = Image.open(img_path)
        else:
            kind = source[0]
            if kind == "intro":
                img = make_intro_slide()
            elif kind == "outro":
                img = make_outro_slide()
            elif kind == "chapter_card":
                img = make_chapter_card(source[1])
            elif kind == "pdf_qa":
                img = pdf_page_to_image(PDF_QA, source[1])
            elif kind == "pdf_cxc":
                img = pdf_page_to_image(PDF_CXC, source[1])
            elif kind == "glossary":
                img = make_glossary_slide(source[1])
            else:
                raise ValueError(f"Unknown source kind: {kind}")
            img.save(img_path)
            print(f"  [image] saved {img_path.name}")

        # ── 2. Generate audio ──────────────────────────────────────────────
        audio_path = AUDIO_DIR / f"{sid}.mp3"
        await generate_audio(narration, audio_path)

        # ── 3. Build clip ──────────────────────────────────────────────────
        clip = build_video_clip(img, audio_path, is_chapter_card=is_chapter_card)
        clips.append(clip)
        total_dur += clip.duration
        print(f"  [clip] duration={clip.duration:.1f}s  cumulative={total_dur:.1f}s")

    print(f"\n[pipeline] Total estimated duration: {total_dur:.1f}s ({total_dur/60:.1f} min)")

    # ── 4. Assemble and render ─────────────────────────────────────────────
    assemble_video(clips, output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate QA TeleVentas training video")
    parser.add_argument(
        "--output", type=Path, default=OUTPUT_MP4,
        help="Output MP4 file path"
    )
    parser.add_argument(
        "--images-only", action="store_true",
        help="Only generate slide images (no audio/video)"
    )
    args = parser.parse_args()

    if args.images_only:
        # Phase 1.5: preview images without audio
        IMAGE_DIR.mkdir(exist_ok=True)
        for slide in SLIDES:
            sid = slide["id"]
            source = slide["source"]
            img_path = IMAGE_DIR / f"{sid}.png"
            if img_path.exists():
                print(f"[skip] {img_path.name} already exists")
                continue
            kind = source[0]
            if kind == "intro":
                img = make_intro_slide()
            elif kind == "outro":
                img = make_outro_slide()
            elif kind == "chapter_card":
                img = make_chapter_card(source[1])
            elif kind == "pdf_qa":
                img = pdf_page_to_image(PDF_QA, source[1])
            elif kind == "pdf_cxc":
                img = pdf_page_to_image(PDF_CXC, source[1])
            elif kind == "glossary":
                img = make_glossary_slide(source[1])
            else:
                continue
            img.save(img_path)
            print(f"[saved] {img_path.name}")
        print(f"\nSlide images saved to: {IMAGE_DIR}")
    else:
        asyncio.run(main(args.output))
