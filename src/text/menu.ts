// MENU
let menuTexts: Record<string, MainMenuText> = {
  es: {
    menu: "Herramientas CoMapeo",
    translateCoMapeoCategory: "Gestionar idiomas y traducir",
    generateIcons: "Generar Íconos para Categorías",
    generateCoMapeoCategory: "Generar Categoría CoMapeo",
    generateCoMapeoCategoryDebug: "Exportar archivos sin procesar",
    debugMenuTitle: "Depuración",
    importCategoryFile: "Importar archivo de categoría",
    importCoMapeoCategory: "Importar archivo de categoría",
    lintAllSheets: "Validar Planillas",
    cleanAllSheets: "Resetear Planillas",
    openHelpPage: "Ayuda",
  },
  en: {
    menu: "CoMapeo Tools",
    translateCoMapeoCategory: "Manage Languages & Translate",
    generateIcons: "Generate Category Icons",
    generateCoMapeoCategory: "Generate CoMapeo Category",
    generateCoMapeoCategoryDebug: "Export Raw Files",
    debugMenuTitle: "Debug",
    importCategoryFile: "Import category file",
    importCoMapeoCategory: "Import category file",
    lintAllSheets: "Lint Sheets",
    cleanAllSheets: "Reset Spreadsheet",
    openHelpPage: "Help",
  },
  pt: {
    menu: "Ferramentas CoMapeo",
    translateCoMapeoCategory: "Gerenciar idiomas e traduzir",
    generateIcons: "Gerar Ícones para Categorias",
    generateCoMapeoCategory: "Gerar Categoria CoMapeo",
    generateCoMapeoCategoryDebug: "Exportar arquivos brutos",
    debugMenuTitle: "Depuração",
    importCategoryFile: "Importar arquivo de categoria",
    importCoMapeoCategory: "Importar arquivo de categoria",
    lintAllSheets: "Validar Planilhas",
    cleanAllSheets: "Resetar Planilhas",
    openHelpPage: "Ajuda",
  },
};

let translateMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Gestionar idiomas y traducir",
    actionText:
      "Esto traducirá todas las celdas vacías en todas las otras columnas de traducción de lenguages. Continuar?",
    completed: "Traducción Completada",
    completedText: "Todas las planillas fueron traducidas con éxito",
    error: "Error",
    errorText: "Ocurrió un error durante la traducción: ",
  },
  en: {
    action: "Manage Languages & Translate",
    actionText:
      "This will translate all empty cells in the other translation language columns. Continue?",
    completed: "Translation Complete",
    completedText: "All sheets have been translated successfully.",
    error: "Error",
    errorText: "An error occurred during translation: ",
  },
  pt: {
    action: "Gerenciar idiomas e traduzir",
    actionText:
      "Isso traduzirá todas as células vazias nas outras colunas de tradução de idiomas. Continuar?",
    completed: "Tradução Concluída",
    completedText: "Todas as planilhas foram traduzidas com sucesso.",
    error: "Erro",
    errorText: "Ocorreu um erro durante a tradução: ",
  },
};

let iconMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Generar Íconos",
    actionText:
      "Esta acción generará íconos usando la información de la planilla actual. Esto puede llevar algunos minutos para processar. ¿Continuar?",
    error: "Error",
    errorText: "Un error ocurrió generando los íconos: ",
  },
  en: {
    action: "Generate Icons",
    actionText:
      "This will generate icons based on the current spreadsheet data. It may take a few minutes to process. Continue?",
    error: "Error",
    errorText: "An error occurred while generating the icons: ",
  },
  pt: {
    action: "Gerar Ícones",
    actionText:
      "Esta ação gerará ícones usando as informações da planilha atual. Isso pode levar alguns minutos para processar. Continuar?",
    error: "Erro",
    errorText: "Ocorreu um erro ao gerar os ícones: ",
  },
};

let categoryMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Generar Categorías de CoMapeo",
    actionText:
      "Esto generará las categorías de CoMapeo basándose en la información de la planilla actual. Puede llevar unos minutos procesar. ¿Continuar?",
    error: "Error",
    errorText: "Ocurrió un error mientras se generaba la categoría: ",
  },
  en: {
    action: "Generate CoMapeo Category",
    actionText:
      "This will generate a CoMapeo category based on the current spreadsheet data. It may take a few minutes to process. Continue?",
    error: "Error",
    errorText: "An error occurred while generating the category: ",
  },
  pt: {
    action: "Gerar Categorias CoMapeo",
    actionText:
      "Isso gerará as categorias do CoMapeo com base nos dados da planilha atual. Pode levar alguns minutos para processar. Continuar?",
    error: "Erro",
    errorText: "Ocorreu um erro ao gerar a categoria: ",
  },
};

let categoryDebugMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Debug: Exportar Archivos Sin Procesar",
    actionText:
      "El modo depuración ejecuta el generador estándar (la exportación rawBuild está deprecada). ¿Continuar?",
    error: "Error",
    errorText:
      "Ocurrió un error mientras se generaba la categoría en modo depuración: ",
  },
  en: {
    action: "Debug: Export Raw Files",
    actionText:
      "Debug mode runs the standard generator (rawBuild export is deprecated). Continue?",
    error: "Error",
    errorText:
      "An error occurred while generating the category in debug mode: ",
  },
  pt: {
    action: "Depuração: Exportar Arquivos Brutos",
    actionText:
      "O modo de depuração executa o gerador padrão (a exportação rawBuild está obsoleta). Continuar?",
    error: "Erro",
    errorText:
      "Ocorreu um erro ao gerar a categoria no modo de depuração: ",
  },
};

let lintMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Validar Categorías de CoMapeo",
    actionText:
      "Esto validará todas las planillas en la hoja de cálculo. ¿Continuar?",
    completed: "Validación terminada",
    completedText: "Todas las planillas fueron validadas con éxito",
    error: "Error",
    errorText: "Un error ocurrió en la validación: ",
  },
  en: {
    action: "Lint CoMapeo Category",
    actionText: "This will lint all sheets in the spreadsheet. Continue?",
    completed: "Linting Complete",
    completedText: "All sheets have been linted successfully.",
    error: "Error",
    errorText: "An error occurred during linting: ",
  },
  pt: {
    action: "Validar Categorias CoMapeo",
    actionText:
      "Isso validará todas as planilhas na planilha. Continuar?",
    completed: "Validação Concluída",
    completedText: "Todas as planilhas foram validadas com sucesso.",
    error: "Erro",
    errorText: "Ocorreu um erro durante a validação: ",
  },
};

let cleanAllMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Resetear Plantillas",
    actionText:
      "!Atención! Esto eliminará todas las traducciones, metadata e íconos the la hoja de cálculos. Esta acción no se puede revertir. ¿Continuar?",
    completed: "Reseteo Completado",
    completedText: "Todas las planillas fueron reseteadas con éxito",
    error: "Error",
    errorText: "Un error ocurrió durante el reseteo: ",
  },
  en: {
    action: "Reset Spreadsheet",
    actionText:
      "Attention! This will remove all translations, metadata, and icons from the spreadsheet. This action cannot be undone. Continue?",
    completed: "Reset Complete",
    completedText: "All sheets have been reset successfully.",
    error: "Error",
    errorText: "An error occurred during reset: ",
  },
  pt: {
    action: "Resetar Planilhas",
    actionText:
      "Atenção! Isso removerá todas as traduções, metadados e ícones da planilha. Esta ação não pode ser desfeita. Continuar?",
    completed: "Redefinição Concluída",
    completedText: "Todas as planilhas foram redefinidas com sucesso.",
    error: "Erro",
    errorText: "Ocorreu um erro durante a redefinição: ",
  },
};

// Alias used by index.ts (v2 API flow)
let importMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Importar archivo de categoría",
    actionText:
      "Esto te permitirá importar un archivo de categoría de CoMapeo (.comapeocat) o archivo de configuración de Mapeo (.mapeosettings) para editar. ADVERTENCIA: Esto borrará todos los datos actuales de la hoja de cálculo y los reemplazará con el contenido del archivo. ¿Continuar?",
    completed: "Importación Completada",
    completedText: "El archivo de categoría ha sido importado con éxito.",
    error: "Error",
    errorText: "Un error ocurrió durante la importación: ",
  },
  en: {
    action: "Import category file",
    actionText:
      "This will allow you to import a CoMapeo category file (.comapeocat) or Mapeo settings file (.mapeosettings) for editing. WARNING: This will erase all current spreadsheet data and replace it with content from the file. Continue?",
    completed: "Import Complete",
    completedText: "The category file has been successfully imported.",
    error: "Error",
    errorText: "An error occurred during import: ",
  },
  pt: {
    action: "Importar arquivo de categoria",
    actionText:
      "Isso permitirá importar um arquivo de categoria CoMapeo (.comapeocat) ou arquivo de configuração Mapeo (.mapeosettings) para edição. AVISO: Isso apagará todos os dados atuais da planilha e os substituirá pelo conteúdo do arquivo. Continuar?",
    completed: "Importação Concluída",
    completedText: "O arquivo de categoria foi importado com sucesso.",
    error: "Erro",
    errorText: "Ocorreu um erro durante a importação: ",
  },
};

// Backward-compatible alias for legacy references (if any)
let importCategoryMenuTexts = importMenuTexts;

let testExtractMenuTexts: Record<string, MenuText> = {
  es: {
    action: "Probar extracción y validación",
    actionText:
      "Esto descargará un archivo de prueba y ejecutará el proceso de extracción y validación para diagnosticar problemas. No se modificarán los datos de la hoja de cálculo. ¿Continuar?",
    completed: "Prueba Completada",
    completedText:
      "La prueba de extracción y validación se ha completado con éxito. Revisa los registros para obtener información detallada.",
    error: "Error",
    errorText: "Un error ocurrió durante la prueba: ",
  },
  en: {
    action: "Test extraction and validation",
    actionText:
      "This will download a test file and run the extraction and validation process to diagnose issues. No spreadsheet data will be modified. Continue?",
    completed: "Test Complete",
    completedText:
      "The extraction and validation test has completed successfully. Check the logs for detailed information.",
    error: "Error",
    errorText: "An error occurred during testing: ",
  },
  pt: {
    action: "Testar extração e validação",
    actionText:
      "Isso fará o download de um arquivo de teste e executará o processo de extração e validação para diagnosticar problemas. Nenhum dado da planilha será modificado. Continuar?",
    completed: "Teste Concluído",
    completedText:
      "O teste de extração e validação foi concluído com sucesso. Verifique os registros para informações detalhadas.",
    error: "Erro",
    errorText: "Ocorreu um erro durante o teste: ",
  },
};
