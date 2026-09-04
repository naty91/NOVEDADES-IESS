
# CENASE - Bitácora de Novedades e Incidencias IESS

App lista para desplegar en Streamlit Community Cloud.

## Archivos
- `main.py`: aplicación.
- `requirements.txt`: dependencias.

## Cómo subir a GitHub y desplegar
1. Crea un repositorio nuevo en GitHub.
2. Sube `main.py` y `requirements.txt` a la raíz del repositorio.
3. Entra a Streamlit Community Cloud.
4. Crea una app nueva y selecciona el repositorio.
5. En "Main file path" indica: `main.py`.
6. Despliega.

## Funciones
- Registro de novedades.
- Fecha efectiva separada de fecha de registro IESS.
- Alertas automáticas:
  - entrada: 15 días;
  - salida / enfermedad / modificación de sueldo: 3 días;
  - amarillo si el retraso está asociado a incidencia IESS;
  - rojo si requiere revisar responsabilidad interna.
- Edición y eliminación.
- Filtros.
- PDF individual.
- Reporte consolidado Excel.
- Reporte consolidado PDF.
- Respaldo ZIP con registros y evidencias.
- Restauración del respaldo.
- Importación de una bitácora Excel existente.
- Configuración de fecha de cierre y reapertura del IESS.
- Resumen por incidencia y responsable.

## Importante sobre persistencia
Streamlit Community Cloud puede reiniciar la app y su disco local no debe tomarse como almacenamiento permanente.
Por eso esta versión usa un respaldo descargable `.zip`.
La práctica recomendada es bajar el respaldo después de cada sesión de trabajo o al cierre del día.

## Privacidad
No se incluyen en este paquete las cédulas ni nombres del archivo real de CENASE.
La información se incorpora desde la propia app o importando el Excel interno.

## Regla operativa central
La fecha efectiva del ingreso/salida debe conservar la realidad laboral. No se modifica para hacerla coincidir con la fecha de disponibilidad del portal del IESS.
