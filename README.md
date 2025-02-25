# 🎙️ Voice2Sheets

**Voice2Sheets** es una herramienta de automatización financiera diseñada para pequeños comerciantes, como dueños de tiendas o misceláneas. Utiliza reconocimiento de voz en español para registrar transacciones de **compra** y **venta** de manera rápida, eficiente y sin intervención manual constante. Los datos se guardan automáticamente en archivos Excel organizados por días y meses.

---

## 🎯 Características principales

- 🎙️ **Reconocimiento de voz en español** para registrar operaciones.
- 📊 **Organización automática de datos** en archivos Excel por días y meses.
- ✅ **Confirmación de cada transacción** antes de guardarla.
- 🛍️ Separación clara de **compras** y **ventas** en cada hoja diaria.
- 💰 **Cálculo automático de totales** diarios de ingresos y egresos.

---

## 🛠️ Tecnologías utilizadas

- **Python 3.12**
- 🎙️ `SpeechRecognition` – Reconocimiento de voz.
- 🔢 `number-parser` – Interpretación de números escritos en palabras.
- 📊 `openpyxl` – Manipulación de archivos Excel (.xlsx).
- 📅 `datetime` – Registro de fecha y hora exacta.
- 🔍 Expresiones regulares (`regex`) – Análisis y extracción de datos del comando de voz.

---

## ⚙️ Instalación

1. Clona este repositorio:
   ```bash
   git clone https://github.com/tuusuario/voice2sheets.git
   cd voice2sheets
