import { GoogleGenAI } from "@google/genai";

export async function generarTexto(prompt, apiKey) {
  console.log("Iniciando generarTexto con el prompt:", prompt);
  
  if (!apiKey) {
    return "Error: No se proporcionó una API Key";
  }

  try {
    // Crear cliente Gemini dinámicamente con la clave ingresada
    const ai = new GoogleGenAI({
      apiKey: apiKey 
    });

    console.log("Llamando a ai.models.generateContent...");
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
    });

    console.log("Respuesta de Gemini recibida:", response);
    return response.text; 

  } catch (error) {
    console.error("Error Gemini:", error);
    return "Error generando contenido";
  }
}