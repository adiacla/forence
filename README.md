# Agente de IA Forense — Versión Académica (Gradio + OpenAI)

Aplicación Gradio que usa la API de OpenAI para apoyar la elaboración de
**borradores** de informes periciales en el flujo de Medicina Legal
(lesiones personales), a partir de documentos cargados por el usuario e
instrucciones adicionales limitadas estrictamente a la redacción del
informe. Permite descargar el resultado como documento **Word (.docx)**.

> ⚠️ Versión académica. No sustituye el criterio ni la firma de un perito
> forense habilitado. No cargue datos reales de pacientes o víctimas.

## 1. Requisitos

- Python 3.10 o superior
- Una API key de OpenAI (https://platform.openai.com/api-keys)

## 2. Instalación local

```bash
# 1. Cree y active un entorno virtual (opcional pero recomendado)
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate

# 2. Instale las dependencias
pip install -r requirements.txt

# 3. Configure su API key como variable de entorno
export OPENAI_API_KEY="sk-..."       # En Windows (PowerShell): $env:OPENAI_API_KEY="sk-..."

# 4. Ejecute la aplicación
python app.py
```

Gradio abrirá la aplicación en `http://127.0.0.1:7860`. También puede pegar
su API key directamente en el campo "Configuración avanzada" de la
interfaz si prefiere no usar una variable de entorno (útil para pruebas
rápidas; no se guarda en ningún archivo).

## 3. Publicar en Hugging Face Spaces ("Gradio Cloud")

El servicio gratuito de hospedaje para apps Gradio es **Hugging Face
Spaces**. Hay dos formas sencillas de publicar esta app:

### Opción A — Comando `gradio deploy` (más rápido)

```bash
pip install -U gradio
gradio deploy
```

Este comando, ejecutado dentro de la carpeta del proyecto, lo guía paso a
paso: crea el Space, sube `app.py` y `requirements.txt`, y le pide
configurar los *secrets* (allí debe agregar `OPENAI_API_KEY` con su clave,
para no exponerla en el código).

### Opción B — Manual desde huggingface.co

1. Cree una cuenta en https://huggingface.co y luego un nuevo **Space**
   (`New Space`), seleccionando **SDK: Gradio**.
2. Suba los archivos `app.py`, `requirements.txt` y este `README.md` al
   repositorio del Space (por interfaz web o `git push`).
3. En la pestaña **Settings → Repository secrets** del Space, agregue un
   secreto llamado `OPENAI_API_KEY` con su clave de OpenAI.
4. El Space se construirá automáticamente y quedará disponible en una URL
   pública tipo `https://huggingface.co/spaces/<usuario>/<nombre-space>`.

> Nunca escriba su API key directamente dentro de `app.py` antes de subirlo
> a un repositorio público. Use siempre el mecanismo de *secrets*.

## 4. Uso de la aplicación

1. Cargue en el panel lateral los documentos del caso (historia clínica,
   matriz de elementos vulnerantes, notas de examen, etc.) en PDF, Word o
   texto plano.
2. Escriba, si lo desea, **instrucciones adicionales sobre el informe**
   (énfasis, formato, secciones extra). El campo está diseñado para
   **rechazar** preguntas generales o temas ajenos al informe: si detecta
   algo fuera de alcance, mostrará un mensaje de rechazo en vez de generar
   el informe.
3. Presione **"Generar informe"**. El borrador aparece en el panel
   derecho, encabezado siempre con el aviso de que es un borrador generado
   con apoyo de IA pendiente de validación pericial.
4. Descargue el resultado como documento **Word (.docx)** con el botón
   correspondiente.
5. Use **"Nueva sesión"** para limpiar todo. La aplicación no guarda
   historial de conversación ni de informes: cada generación es una
   llamada independiente al modelo, sin memoria de turnos anteriores ni
   persistencia entre sesiones.

## 5. Estructura del proyecto

```
.
├── app.py              # Aplicación Gradio (interfaz + lógica + llamada a OpenAI + export a Word)
├── requirements.txt     # Dependencias
└── README.md            # Este archivo
```

## 6. Limitaciones y alcance (léase)

- El agente está instruido para **no responder preguntas generales**, no
  sostener conversación y no revelar sus instrucciones de sistema; solo
  redacta o ajusta el borrador del informe pericial. Además del control por
  instrucciones (system prompt), la app incluye un filtro adicional en
  Python que bloquea de forma temprana instrucciones claramente fuera de
  alcance, antes de llamar a la API.
- Ningún filtro automático es infalible. La responsabilidad final sobre el
  contenido del informe y su uso recae siempre en el perito o docente a
  cargo.
- No se debe cargar información real de pacientes o víctimas; use datos
  sintéticos o anonimizados con fines académicos.
