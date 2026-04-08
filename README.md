# 🛡️ Syllabus Virtualizer TGA - ESUFA

¡Bienvenido al **Syllabus Virtualizer** para la Tecnología en Gestión de Recursos Aeronáuticos (TGA)! Esta herramienta ha sido diseñada para modernizar y agilizar el proceso de creación de sílabos institucionales, integrando inteligencia artificial avanzada para garantizar el alineamiento constructivo y la calidad pedagógica.

## 🚀 ¿Qué es esta herramienta?
Es una aplicación web interactiva que permite a los docentes de la **ESUFA**:
*   Seleccionar asignaturas existentes de la malla TGA.
*   Diseñar Objetivos y Resultados de Aprendizaje (RAPs) precisos.
*   Estructurar unidades temáticas para modalidades Presenciales y Virtuales (AVA/Blackboard).
*   **Magia IA:** Utilizar el cerebro de Google Gemini para alinear automáticamente temas y evaluaciones con los RAPs bajo la taxonomía de Bloom.
*   Exportar el sílabo a formato PDF listo para gestión institucional.

## 🧪 Configuración de la IA (Paso Vital)
Para que las funciones de alineamiento y optimización funcionen, debe configurar la conexión con la IA:

1.  Haga clic en el botón ⚙️ **IA Settings** en la parte superior.
2.  Ingrese la **URL del Worker/Proxy** (proporcionado por el administrador).
3.  Opcionalmente, ingrese su propia **API Key de Gemini** si desea usar una conexión directa personal.
4.  Seleccione el modelo (se recomienda **Gemini Flash** por su velocidad y estabilidad).

> **Nota:** El sistema cuenta con **reintentos automáticos**. Si los servidores de Google están saturados, la herramienta esperará y lo intentará de nuevo por usted.

## 🎓 Enfoque Pedagógico
Esta herramienta prohíba estrictamente el uso de metodologías pasivas ("Clase Magistral"). Todo el contenido generado por la IA está alineado con:
*   **Aprendizaje Activo:** Fomento de la participación y resolución de problemas.
*   **Virtualización:** Estructura optimizada para ambientes virtuales de aprendizaje.
*   **Alineamiento Constructivo:** Metodología de G. Biggs (RAPs -> Actividades -> Evaluación).

## 🛡️ Seguridad y Privacidad
*   **Sin claves expuestas:** Su API Key nunca se guarda en el código público. Se almacena de forma segura en su navegador o se gestiona a través de un proxy cifrado.
*   **Historial Local:** Sus sílabos se guardan en el historial del navegador (`localStorage`), garantizando que su trabajo no se pierda.

## 📂 Estructura del Proyecto
*   `index.html`: Interfaz de usuario.
*   `style.css`: Estilos visuales institucionales.
*   `script.js`: Cerebro y lógica de integración con la IA.
*   `tga_syllabus_db.json`: Base de datos de la malla curricular.

---
**Desarrollado para el fortalecimiento académico de la Escuela de Suboficiales de la Fuerza Aeroespacial Colombiana (ESUFA).** 🛰️💪
