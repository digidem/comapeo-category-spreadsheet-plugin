declare const google: any;

declare const VERSION: string;
declare const COMMIT: string;
declare const BRANCH: string;

declare function getVersionInfo(): string;
declare function getFullVersionInfo(): string;

declare function testLanguageLookup(): void;
declare function testLanguageLookupQuick(): void;
declare function testLanguageRecognitionIntegration(): void;
declare function testUtilsSlugify(): void;
declare function testFormatDetection(url?: string): unknown;
declare function testFieldExtraction(url?: string): unknown;
declare function testExtractAndValidate(): void;
declare function testDetailsAndIcons(): void;
declare function testTranslationExtraction(url?: string): unknown;
declare function testImportCategory(url?: string): unknown;
declare function testZipToApi(): void;
declare function testEndToEnd(url?: string): unknown;
declare function testSkipTranslation(): void;
declare function testDebugLogger(): void;
