import pandas as pd
from reportlab.lib.pagesizes import letter
from io import BytesIO
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

# Function to add a new page to the document
def new_page(doc, elements):
    doc.build(elements)
    elements = []
    return SimpleDocTemplate(BytesIO(), pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30), elements

# Load geoMap data
geoMap_dataset = pd.read_csv('geoMap.csv', names=['Country', 'Search Volume'], header=None)
geoMap_dataset = geoMap_dataset.dropna()

# Get top 25 searches in geoMap
top_geoMap = geoMap_dataset.nlargest(25, 'Search Volume')

# Calculate the width of the graph (100% of the page - 10px margin)
max_width = letter[0] - 10

# Create SimpleDocTemplate object with margins
pdf_buffer = BytesIO()
pdf = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)

# List to store elements
elements = []

# Style for the title
title_style = ParagraphStyle('Title', parent=getSampleStyleSheet()['Title'], spaceAfter=12)

# Style for the subtitle
subtitle_style = ParagraphStyle('Subtitle', parent=getSampleStyleSheet()['Heading2'], spaceAfter=6)

# Style for the body text
body_text_style = ParagraphStyle('BodyText', parent=getSampleStyleSheet()['BodyText'], spaceAfter=6)

# Style for the conclusive text
conclusive_text_style = ParagraphStyle('ConclusiveText', parent=getSampleStyleSheet()['BodyText'], spaceAfter=12)

# Title of the report
title = Paragraph("<u>Informe Ejecutivo: Búsquedas de Ciberseguridad</u>", title_style)
elements.append(title)

# Introduction
introduction_text = (
    "Este informe ejecutivo tiene como objetivo analizar las tendencias globales de búsquedas relacionadas con el término 'ciberseguridad'. "
    "La ciberseguridad se ha convertido en un tema crucial en la era digital, y entender las dinámicas de búsqueda proporciona insights "
    "sobre la conciencia y preocupaciones de las personas en diferentes partes del mundo. A través de este análisis, buscamos identificar "
    "patrones, cambios y áreas de interés emergentes en el ámbito de la ciberseguridad."
)
introduction_paragraph = Paragraph(introduction_text, body_text_style)
elements.append(introduction_paragraph)

# Add bar chart (Top 25)
fig, ax = plt.subplots(figsize=(max_width / 80, 8))
ax.bar(top_geoMap['Country'], top_geoMap['Search Volume'], color='blue')
ax.set_xlabel('Country')
ax.set_ylabel('Search Volume')

# Title of the chart
graph_title = Paragraph("<u>Búsquedas de Ciberseguridad por País (Top 25)</u>", subtitle_style)
elements.append(graph_title)

# Rotate x-axis labels to maintain diagonal arrangement
plt.xticks(rotation=45, ha='right')

# Save the chart to BytesIO
geoMap_buffer = BytesIO()
canvas = FigureCanvas(fig)
canvas.print_png(geoMap_buffer)
plt.close()

# Add chart as an image
elements.append(Spacer(1, 20))
geoMap_buffer.seek(0)
elements.append(Image(geoMap_buffer, width=letter[0], height=400))
elements.append(Spacer(1, 10))

# Summary
summary_text = (
    "En resumen, valorando la tendencia global de búsquedas de ciberseguridad, observamos un aumento constante en el interés "
    "y la conciencia sobre este tema. Los datos muestran que las personas en diferentes países buscan activamente información "
    "relacionada con la ciberseguridad, reflejando la creciente importancia de este campo. Esta tendencia subraya la necesidad "
    "de continuar fortaleciendo las medidas de ciberseguridad a nivel mundial para abordar las amenazas digitales en evolución."
)
summary_paragraph = Paragraph(summary_text, conclusive_text_style)
elements.append(summary_paragraph)

# Build the PDF
pdf, elements = new_page(pdf, elements)

# Save the PDF to a file
pdf_buffer.seek(0)
with open("Cybersecurity_Report.pdf", "wb") as f:
    f.write(pdf_buffer.read())
