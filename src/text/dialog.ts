let selectTranslationLanguagesDialogText: Record<string, SelectTranslationDialogText> = {
  es: {
    title: "Seleccionar idiomas de destino",
    buttonText: "Traducir",
    skipButtonText: "Cancelar",
    message: [
      "Selecciona los idiomas a los que deseas traducir automáticamente desde {{sourceLanguage}}:",
      "Las traducciones se aplicarán solo a los idiomas marcados abajo.",
      "Si quieres trabajar con nuevos idiomas manualmente, agrégalos en la sección inferior.",
    ],
    manualSectionTitle: "Idiomas personalizados (sin traducción automática)",
    manualSectionDescription: [
      "Agrega columnas para idiomas adicionales en todas las hojas de traducción.",
      "Estos idiomas no se traducen automáticamente; podrás completarlos manualmente más tarde.",
    ],
    manualDropdownPlaceholder: "Selecciona un idioma común",
    manualAddButton: "Agregar otro idioma",
    manualNamePlaceholder: "Nombre del idioma (ej. Quechua)",
    manualIsoPlaceholder: "Código ISO (ej. qu)",
    validationMessages: {
      noAutoSelection: "Selecciona al menos un idioma para traducir o haz clic en 'Cancelar'.",
      missingCustomFields: "Completa el nombre y el código ISO para cada idioma personalizado.",
      duplicateCustomIso: "Hay códigos ISO personalizados duplicados. Verifica e intenta nuevamente.",
      invalidCustomIso: "Algunos códigos ISO personalizados no tienen formato válido. Usa letras (a-z) y guiones bajos.",
    },
  },
  en: {
    title: "Select Target Languages",
    buttonText: "Translate",
    skipButtonText: "Cancel",
    message: [
      "Choose the languages you want to auto-translate from {{sourceLanguage}}:",
      "Only the languages you select below will receive automatic translations.",
      "Add any manual-only languages in the section underneath.",
    ],
    manualSectionTitle: "Custom languages (manual columns)",
    manualSectionDescription: [
      "Add additional language columns across every translation sheet.",
      "These languages are not auto-translated—you can fill them in later.",
    ],
    manualDropdownPlaceholder: "Select a common language",
    manualAddButton: "Add another language",
    manualNamePlaceholder: "Language name (e.g. Quechua)",
    manualIsoPlaceholder: "ISO code (e.g. qu)",
    validationMessages: {
      noAutoSelection: "Select at least one language to translate or click 'Cancel'.",
      missingCustomFields: "Please fill in both the language name and ISO code for each custom language.",
      duplicateCustomIso: "Duplicate custom ISO codes detected. Please ensure each custom language has a unique code.",
      invalidCustomIso: "Custom ISO codes should contain letters and optional hyphens (e.g. qu, quz, pt-br).",
    },
  },
};

let iconDialogTexts: Record<string, DialogText> = {
  es: {
    title: "Íconos de CoMapeo Generados",
    message: [
      "Tus íconos de CoMapeo fueron creados con éxito y guardados a una carpeta en tu Google Drive.",
      "Para ver y administrar tus íconos generados, haz click en el botón debajo. Puedes descargar, modificar o reemplazar los íconos como necesites.",
      "Recuerda actualizar las URLs en la planillas si haces algún cambio.",
    ],
    buttonText: "Ver íconos generados",
  },
  en: {
    title: "CoMapeo Icons Generated",
    message: [
      "Your CoMapeo icons have been successfully generated and saved to a folder in your Google Drive.",
      "To view and manage your generated icons, click the button below. You can download, modify, or replace icons as needed.",
      "Remember to update the icon URLs in the spreadsheet if you make any changes.",
    ],
    buttonText: "View Generated Icons",
  },
};

let generatedConfigDialogTexts: Record<string, DialogText> = {
  es: {
    title: "Categorías de CoMapeo generadas",
    message: [
      "Tu archivo de Categoría CoMapeo fue generado con éxito y guardado en Google Drive.",
      "Para descargarlo o compartirlo, haz click en el botón de abajo. El archivo con extensión '.comapeocat' está listo para ser importado en la app de CoMapeo.",
    ],
    buttonText: "Descargar categorías de CoMapeo",
  },
  en: {
    title: "CoMapeo Category Generated",
    message: [
      "Your CoMapeo Category file has been successfully generated and saved to Google Drive.",
      "To download or share it, click the button below. The .comapeocat file is ready to import into the CoMapeo app.",
    ],
    buttonText: "Download CoMapeo Category",
  },
};

let helpDialogTexts: Record<string, DialogText & DialogInstructions> = {
  es: {
    title: "Herramientas de CoMapeo - Ayuda",
    message: [
      "Bienvenido a las herramientas de CoMapeo! Este add-on te ayudará a organizar y crear categorías de CoMapeo. Se usa de esta manera:",
      "La lógica de trabajo para crear y organizar categorías de CoMapeo es la siguiente:",
    ],
    instructions: [
      // TODO: sheets shouls also be translated
      "Edita las hojas llamadas 'Categories' y 'Details' para definir sus categorías propias y sus detalles asociados. Note que el color de fondo que seleccione para las categorías e íconos se verá reflejado en la app de CoMapeo",
      "Usa la opción 'Gestionar idiomas y traducir' para agregar idiomas personalizados y generar traducciones automáticamente para las celdas vacías",
      "Revise y refine las traducciones auto-generadas cuánto sea necesario",
      "Use la opción 'Generar Íconos para Categorías' para crear íconos para sus categorías. El color de fondo de los íconos coincidirá con el color que eliga en la planilla.",
      "Revise los íconos generados en la caerpa de íconos y modífiquelos usando la <br /><a href='https://icons.earthdefenderstoolkit.com' target='_blank'>Icon Generator App</a>  de ser necessario",
      "Copie el link compartido para cada ícono y péguelo en la celda de ícono correspondiente en la planilla",
      "Use la opción 'Validar Planillas' para asegurarse que su información está en el formato y capitalización correcta.",
      "Repita los pasos anteriores tal como necesite, actualizando traducciones e íconos hasta que esté conforme con los resultados.",
      "Cuando sus categorías estén listas, use la opcíon 'Generar Categoría CoMapeo' para crear su archivo final. Este proceso puede tardar varios minutos y generará un archivo '.comapeocat' guardado en Google Drive, listo para ser importado en la app de CoMapeo.",
    ],
    footer: "Para más información, visite nuestro repositorio de Github",
    buttonText: "Visite repositorio de Github",
  },
  en: {
    title: "CoMapeo Tools Help",
    message: [
      "Welcome to CoMapeo Tools! This add-on helps you manage and generate CoMapeo categories. Here's how to use it:",
      "The general workflow for creating and managing CoMapeo categories is as follows:",
    ],
    instructions: [
      "Edit the 'Categories' and 'Details' sheets to define your custom categories and their associated details. Note that the background color you set for categories and icons will reflect their color in the CoMapeo app.",
      "Use the 'Manage Languages & Translate' option to add manual language columns and automatically translate empty cells in other language columns.",
      "Review and refine the auto-generated translations as needed.",
      "Use the 'Generate Icons' option to create icons for your categories. The background color of the icons will match the color you set in the spreadsheet.",
      "Check the generated icons in the icons folder and modify them using the <br /><a href='https://icons.earthdefenderstoolkit.com' target='_blank'>Icon Generator App</a> if necessary.",
      "Copy the shared link for each icon and paste it into the corresponding icon cell in the spreadsheet.",
      "Use the 'Lint Sheets' option to ensure proper formatting and capitalization of your data.",
      "Repeat the previous steps as needed, updating translations and icons until you're satisfied with the results.",
      "When ready, use the 'Generate CoMapeo Category' option to create your final file. This process may take a few minutes and will generate a '.comapeocat' file saved to Google Drive, ready for use with the CoMapeo app.",
    ],
    footer: "For more information, visit our Github repository",
    buttonText: "Visit GitHub Repository",
  },
};

let processingDialogTitle = {
  en: "Generating CoMapeo Category",
  es: "Generando Categoría CoMapeo",
};

let iconErrorDialogTexts: Record<string, IconErrorDialogText> = {
  en: {
    title: "Icon Processing Report",
    downloadButtonText: "Download Error Report",
    continueButtonText: "Continue Anyway",
    okButtonText: "OK",
  },
  es: {
    title: "Reporte de Procesamiento de Íconos",
    downloadButtonText: "Descargar Reporte de Errores",
    continueButtonText: "Continuar de Todos Modos",
    okButtonText: "OK",
  },
};

let processingDialogTexts: Record<string, DialogText>[] = [
  // Step 1: Initializing (merges old steps 1-4)
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Initializing... (1/8)", "Validating and reading spreadsheet data"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Inicializando... (1/8)", "Validando y leyendo datos de la planilla"],
    },
  },
  // Step 2: Translating (conditional - only shown if languages selected)
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Translating... (2/8)", "Auto-translating to selected languages"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Traduciendo... (2/8)", "Generando traducciones automáticas"],
    },
  },
  // Step 3: Processing data
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Preparing build request... (3/8)", "Collecting spreadsheet data"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Preparando solicitud... (3/8)", "Recopilando datos de la planilla"],
    },
  },
  // Step 4: Building payload
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Building payload... (4/8)", "Creating JSON for the API"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Creando payload... (4/8)", "Generando JSON para la API"],
    },
  },
  // Step 5: Sending request
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Sending request... (5/8)", "Submitting build to API"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Enviando solicitud... (5/8)", "Enviando generación a la API"],
    },
  },
  // Step 6: Waiting for response
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Waiting for response... (6/8)", "This may take a few minutes"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Esperando respuesta... (6/8)", "Esto puede tardar algunos minutos"],
    },
  },
  // Step 7: Saving to Drive
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Saving to Drive... (7/8)", "Storing the .comapeocat file"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["Guardando en Drive... (7/8)", "Guardando el archivo .comapeocat"],
    },
  },
  // Step 8: Complete
  {
    en: {
      title: processingDialogTitle["en"],
      message: ["Complete! (8/8)", "CoMapeo Category ready for download"],
    },
    es: {
      title: processingDialogTitle["es"],
      message: ["¡Completo! (8/8)", "Categoría CoMapeo lista para descargar"],
    },
  },
];
