import { insertarCuerpo, insertarEncabezado } from "./template";
import { reemplazarPlaceholders } from "./placeholders";
import { generarTexto } from "./ia";

Office.onReady(() => {

  document
    .getElementById("insertar-encabezado")
    .onclick = insertarEncabezado;

  document
    .getElementById("insertar-cuerpo")
    .onclick = insertarCuerpo;


  document
  .getElementById("generar-intro")
  .onclick = async () => {
    console.log("Botón 'Generar PDC' clickeado.");
    
    const btn = document.getElementById("generar-intro");
    const loader = document.getElementById("ai-loader");
    const btnText = btn.querySelector(".btn-text");
    
    // Cambiar estado visual a cargando
    btn.classList.add("btn-disabled");
    loader.style.display = "block";
    btnText.innerText = "Generando...";

    const apiKey = document.getElementById("api-key-input").value.trim();
    if (!apiKey) {
      console.error("API Key vacía. Por favor ingrese una clave válida.");
      return; // Detiene la ejecución si no hay clave
    }

    const materia = document.getElementById("materia-input").value.trim();
    const curso = document.getElementById("curso-input").value.trim();
    const temario = document.getElementById("temario-input").value.trim();

    try {
      const promptText = `Eres un EXPERTO en diseño curricular boliviano bajo el Modelo Educativo Sociocomunitario Productivo (Ley 070 de la Educación "Avelino Siñani - Elizardo Pérez").

DATOS DEL PDC:
- Materia: ${materia}
- Curso: ${curso}
- Contenidos: ${temario}

===============================================================================
ESTRUCTURA OBLIGATORIA DEL PDC (4 SEMANAS)
===============================================================================

SEMANA 1: Introducción y bases conceptuales (primeros 2-3 temas)
SEMANA 2: Desarrollo y profundización (siguientes 2-3 temas)
SEMANA 3: Aplicación y consolidación (siguientes temas)
SEMANA 4: Evaluación, producción y cierre integrador

===============================================================================
REGLAS DE REDACCIÓN OBLIGATORIAS
===============================================================================

1. OBJETIVO DE APRENDIZAJE:
✓ Redactar en modo indicativo y primera persona plural.
✓ Integrar dimensiones SER, SABER, HACER y DECIDIR.
✓ Debe relacionarse con todos los contenidos del trimestre.
✓ Extensión: 2-3 líneas continuas.

2. CONTENIDOS POR SEMANA:
✓ Distribuir equilibradamente los contenidos.
✓ Redactar de forma concisa y coherente.
✓ Semana 4 debe incluir evaluación integral y producción final.

3. PRÁCTICA (SEMANA 1):
✓ Redactar en primera persona plural.
✓ Actividades vivenciales y contextualizadas.
✓ Relacionadas con la realidad sociocomunitaria.

4. TEORÍA (SEMANA 1 y 2):
✓ Redactar en primera persona plural.
✓ Sistematizar conceptos, principios y contenidos.
✓ Relacionar teoría con práctica.

5. VALORACIÓN (SEMANA 3 y 4):
✓ Reflexión crítica y comunitaria.
✓ Relacionar con valores sociocomunitarios y Vivir Bien.
✓ Incluir obligatoriamente:
- ¿Qué fue lo más fácil de comprender?
- ¿Qué fue lo más difícil de aprender?

6. PRODUCCIÓN (SEMANA 3 y 4):
✓ Producto tangible y aplicable.
✓ Semana 4 debe integrar todos los aprendizajes del trimestre.

7. RECURSOS:
✓ Solo recursos específicos y especializados.
✓ No incluir materiales básicos repetitivos.

8. PERIODOS:
✓ Todas las semanas deben tener 6 periodos.

9. CRITERIOS:
✓ SER: valores y actitudes.
✓ SABER: conocimientos y comprensión.
✓ HACER: aplicación y procedimientos.
✓ Redactar de forma breve y directa.

⚠️⚠️⚠️ MUY IMPORTANTE - FORMATO DE CRITERIOS ⚠️⚠️⚠️

   Los CRITERIOS deben ser frases DIRECTAS y CONCISAS.

   NO INCLUIR NUNCA estos prefijos:
   ❌ "SABER - Criterio 1:"
   ❌ "HACER - Criterio 2:"
   ❌ "Criterio 1 de SABER:"
   ❌ "Criterio 1:"

   FORMATO CORRECTO (SOLO la frase descriptiva):
   ✅ "Caracterización de compuestos orgánicos e inorgánicos"
   ✅ "Comprensión de la estructura y enlaces del carbono"
   ✅ "Aplicación de técnicas de modelado molecular"
   ✅ "Elaboración de proyectos de química orgánica aplicada"


===============================================================================
PLACEHOLDERS OBLIGATORIOS
===============================================================================

Debes generar el contenido EXACTAMENTE usando estos placeholders:

{{CURSO}}
{{OBJETIVO}}
{{CONTSEM1}}
{{PRACTICASEM1}}
{{TEORIASEM1}}
{{RECURSOSSEM1}}
{{PERIODOS}}
{{CRITERIOSER1}}
{{CRITERIOSER2}}
{{CRITERIOSABER1}}
{{CRITERIOSABER2}}
{{CRITERIOHACER1}}
{{CRITERIOHACER2}}
{{CONTSEM2}}
{{TEORIASEM2}}
{{RECURSOSSEM2}}
{{CONTSEM3}}
{{VALORACIONSEM3}}
{{PRODUCCIONSEM3}}
{{RECURSOSSEM3}}
{{CONTSEM4}}
{{VALORACIONSEM4}}
{{PRODUCCIONSEM4}}
{{RECURSOSSEM4}}

===============================================================================
FORMATO DE RESPUESTA OBLIGATORIO
===============================================================================

Responde ÚNICAMENTE con un objeto JSON válido.

{
  "{{CURSO}}": "grado o curso correspondiente",

  "{{OBJETIVO}}": "objetivo holístico completo",

  "{{CONTSEM1}}": "contenidos semana 1",
  "{{PRACTICASEM1}}": "actividades prácticas",
  "{{TEORIASEM1}}": "desarrollo teórico",
  "{{RECURSOSSEM1}}": "recursos específicos",
  "{{PERIODOS}}": "6",

  "{{CRITERIOSER1}}": "criterio ser 1",
  "{{CRITERIOSER2}}": "criterio ser 2",

  "{{CRITERIOSABER1}}": "criterio saber 1",
  "{{CRITERIOSABER2}}": "criterio saber 2",

  "{{CRITERIOHACER1}}": "criterio hacer 1",
  "{{CRITERIOHACER2}}": "criterio hacer 2",

  "{{CONTSEM2}}": "contenidos semana 2",
  "{{TEORIASEM2}}": "desarrollo teórico semana 2",
  "{{RECURSOSSEM2}}": "recursos específicos semana 2",

  "{{CONTSEM3}}": "contenidos semana 3",
  "{{VALORACIONSEM3}}": "valoración y reflexión",
  "{{PRODUCCIONSEM3}}": "producto tangible",
  "{{RECURSOSSEM3}}": "recursos específicos semana 3",

  "{{CONTSEM4}}": "evaluación integral y producción final",
  "{{VALORACIONSEM4}}": "reflexión final",
  "{{PRODUCCIONSEM4}}": "proyecto integrador",
  "{{RECURSOSSEM4}}": "recursos específicos semana 4"
}

GENERA AHORA EL PDC COMPLETO EN FORMATO JSON.`;

      let iaResponse = await generarTexto(promptText, apiKey);
      console.log("Texto generado (bruto):", iaResponse);

      // Limpiamos los backticks de markdown que suele agregar la IA
      iaResponse = iaResponse.replace(/```json/g, "").replace(/```/g, "").trim();

      const datosJSON = JSON.parse(iaResponse);

      await reemplazarPlaceholders(datosJSON);
      console.log("Placeholders reemplazados con éxito.");
    } catch (e) {
      console.error("Error en el evento onclick o al parsear la respuesta JSON:", e);
    } finally {
      // Restaurar estado visual del botón al terminar
      btn.classList.remove("btn-disabled");
      loader.style.display = "none";
      btnText.innerText = "Generar PDC";
    }
  };

});