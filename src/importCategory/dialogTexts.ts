/**
 * Dialog texts for the import category functionality.
 * This file contains the texts used in the import category dialog.
 */

// Dialog texts for different languages
const importCategoryDialogTexts: Record<string, { title: string; message: string[]; buttonText: string; progressStages: { extracting: string; parsing: string; icons: string; applying: string; finalizing: string }; successMessage: { title: string; dropdownInfo: string[] } }> = {
  en: {
    title: "Import Category File",
    message: [
      "Upload a .comapeocat or .mapeosettings file to import categories into this spreadsheet.",
      "This will replace all current categories, details, and translations.",
      "Make sure you have a backup of your current data before proceeding.",
    ],
    buttonText: "Select File",
    progressStages: {
      extracting: "Extracting files",
      parsing: "Processing category data",
      icons: "Extracting icons",
      applying: "Updating spreadsheet",
      finalizing: "Finalizing",
    },
    successMessage: {
      title: "Import successful!",
      dropdownInfo: [
        "<strong>Details Column Multi-Select:</strong>",
        "The Details column (C) in Categories uses a dropdown with comma-separated values (e.g., 'Name, Type, Notes').",
        "",
        "<strong>Optional:</strong> To enable native multi-select chip display:",
        "1. Select the Details column (C) in Categories sheet",
        "2. Go to Data → Data validation",
        "3. Change 'Display style' from 'Arrow' to 'Chip'",
        "4. Check 'Allow multiple selections'",
        "",
        "Note: This feature is not yet available via Apps Script API.",
      ],
    },
  },
  es: {
    title: "Importar archivo de categoría",
    message: [
      "Sube un archivo .comapeocat o .mapeosettings para importar categorías a esta planilla.",
      "Esto reemplazará todas las categorías, detalles y traducciones actuales.",
      "Asegúrate de tener una copia de seguridad de tus datos actuales antes de continuar.",
    ],
    buttonText: "Seleccionar Archivo",
    progressStages: {
      extracting: "Extrayendo archivos",
      parsing: "Procesando datos de categorías",
      icons: "Extrayendo iconos",
      applying: "Actualizando hoja de cálculo",
      finalizing: "Finalizando",
    },
    successMessage: {
      title: "¡Importación exitosa!",
      dropdownInfo: [
        "<strong>Selección Múltiple en Columna Detalles:</strong>",
        "La columna Detalles (C) en Categorías usa un menú desplegable con valores separados por comas (ej., 'Nombre, Tipo, Notas').",
        "",
        "<strong>Opcional:</strong> Para habilitar la visualización nativa de chips de selección múltiple:",
        "1. Selecciona la columna Detalles (C) en la hoja Categorías",
        "2. Ve a Datos → Validación de datos",
        "3. Cambia 'Estilo de visualización' de 'Flecha' a 'Chip'",
        "4. Marca 'Permitir selecciones múltiples'",
        "",
        "Nota: Esta función aún no está disponible a través de la API de Apps Script.",
      ],
    },
  },
  pt: {
    title: "Importar arquivo de categoria",
    message: [
      "Carregue um arquivo .comapeocat ou .mapeosettings para importar categorias nesta planilha.",
      "Isso substituirá todas as categorias, detalhes e traduções atuais.",
      "Certifique-se de ter um backup dos seus dados atuais antes de prosseguir.",
    ],
    buttonText: "Selecionar Arquivo",
    progressStages: {
      extracting: "Extraindo arquivos",
      parsing: "Processando dados de categoria",
      icons: "Extraindo ícones",
      applying: "Atualizando planilha",
      finalizing: "Finalizando",
    },
    successMessage: {
      title: "Importação bem-sucedida!",
      dropdownInfo: [
        "<strong>Seleção Múltipla na Coluna Detalhes:</strong>",
        "A coluna Detalhes (C) em Categorias usa um menu suspenso com valores separados por vírgula (ex., 'Nome, Tipo, Notas').",
        "",
        "<strong>Opcional:</strong> Para habilitar a exibição nativa de chips de seleção múltipla:",
        "1. Selecione a coluna Detalhes (C) na planilha Categorias",
        "2. Vá em Dados → Validação de dados",
        "3. Mude 'Estilo de exibição' de 'Seta' para 'Chip'",
        "4. Marque 'Permitir seleções múltiplas'",
        "",
        "Nota: Este recurso ainda não está disponível através da API do Apps Script.",
      ],
    },
  },
};
