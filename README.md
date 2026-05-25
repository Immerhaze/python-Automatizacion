Automatización Excel → Formulario 🤖

🇪🇸 Español | 🇺🇸 English below


🇪🇸 Español
Automatización con Puppeteer que eliminó la carga manual de datos desde Excel a un sistema de gestión — reduciendo el tiempo de procesamiento de ~30 segundos por registro a menos de 1 segundo, para más de 100 personas.
El problema
Un trabajador debía transferir manualmente los datos de un archivo Excel a un formulario web — persona por persona. Cada registro implicaba abrir el formulario, rellenar varios campos y enviarlo. Para 100 personas, eso equivalía a más de 50 minutos de trabajo repetitivo.
La solución
Un script con Puppeteer que lee cada fila del Excel y rellena y envía el formulario automáticamente, sin intervención humana.
Resultados
AntesDespuésTiempo por registro~30 segundos< 1 segundo100 registros~50 minutos< 2 minutosEsfuerzo humanoAltoNinguno
Stack

Python
Puppeteer (via Pyppeteer)
openpyxl — lectura del Excel

Correr localmente
bashgit clone https://github.com/Immerhaze/python-Automatizacion
cd python-Automatizacion
pip install -r requirements.txt
python main.py

🇺🇸 English <a name="english"></a>
Puppeteer-based automation that eliminated manual data entry from an Excel spreadsheet into a third-party management system — reducing processing time from ~30 seconds per record to under 1 second, across 100+ entries.
The Problem
A staff member had to manually transfer data from an Excel file into a web-based form — one person at a time. Each record required opening the form, filling multiple fields, and submitting. For 100 people, that's 50+ minutes of repetitive work.
The Solution
A Puppeteer script that reads each row from the Excel file and automatically fills and submits the form, handling the full flow without human intervention.
Results
BeforeAfterTime per record~30 seconds< 1 second100 records~50 minutes< 2 minutesHuman effortHighNone
Tech Stack

Python
Puppeteer (via Pyppeteer)
openpyxl — Excel parsing

Run locally
bashgit clone https://github.com/Immerhaze/python-Automatizacion
cd python-Automatizacion
pip install -r requirements.txt
python main.py

Built by Nicode
