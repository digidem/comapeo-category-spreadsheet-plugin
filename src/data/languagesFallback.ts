/**
 * Fallback language data for when remote fetch fails
 *
 * This data is used as a last resort when:
 * 1. Cache is empty/expired
 * 2. Remote fetch from GitHub fails
 *
 * Enhanced version with both English and native language names to support
 * dual-name recognition (e.g., both "Portuguese" and "Português").
 *
 * ## Data Source
 * Source: https://raw.githubusercontent.com/digidem/comapeo-mobile/refs/heads/develop/src/frontend/languages.json
 * Last updated: 2025-10-28
 * Last verified: 2025-10-28
 * Languages count: 142
 *
 * ## Update Procedure
 *
 * To update this fallback data:
 *
 * 1. **Fetch Latest Data**:
 *    ```bash
 *    curl https://raw.githubusercontent.com/digidem/comapeo-mobile/refs/heads/develop/src/frontend/languages.json > languages.json
 *    ```
 *
 * 2. **Convert to Enhanced Format**:
 *    - Ensure each entry has both `englishName` and `nativeName` fields
 *    - Validate native names are accurate (check with native speakers if possible)
 *    - Keep alphabetical order by language code for maintainability
 *
 * 3. **Update This File**:
 *    - Replace `LANGUAGES_FALLBACK_ENHANCED` object
 *    - Update "Last updated" and "Last verified" dates
 *    - Update "Languages count" if changed
 *
 * 4. **Validation** (recommended):
 *    ```typescript
 *    // In Apps Script editor or test file
 *    const enhanced = LANGUAGES_FALLBACK_ENHANCED;
 *    const lookup = createLanguageLookup(enhanced);
 *
 *    // Verify no collisions
 *    Logger.log(`Total languages: ${lookup.getAllCodes().length}`);
 *
 *    // Test a few samples
 *    Logger.log(lookup.getCodeByName("Portuguese")); // Should be "pt"
 *    Logger.log(lookup.getCodeByName("Português"));  // Should be "pt"
 *    ```
 *
 * 5. **Test Integration**:
 *    - Run `testLanguageLookup()` to verify lookup functionality
 *    - Run `testLanguageRecognitionIntegration()` for end-to-end validation
 *    - Test with actual spreadsheet using various language names
 *
 * ## Verification Checklist
 *
 * - [ ] Native names are accurate for all languages
 * - [ ] No name collisions (different languages with same normalized name)
 * - [ ] All major world languages included
 * - [ ] Consistent formatting (capitalization, spacing)
 * - [ ] Tests pass with updated data
 * - [ ] Documentation dates updated
 *
 * ## Notes
 *
 * - This file is committed to git and serves as last-resort fallback
 * - Remote data is preferred and cached for 6 hours
 * - Changes here affect users only when remote fetch fails
 * - Consider impact on existing spreadsheets when modifying language codes
 */

/// <reference path="../types.ts" />

/**
 * Enhanced fallback language data with dual-name support
 * Maps language codes to both English and native names
 */
const LANGUAGES_FALLBACK_ENHANCED: LanguageMapEnhanced = {
  en: { englishName: "English", nativeName: "English" },
  es: { englishName: "Spanish", nativeName: "Español" },
  pt: { englishName: "Portuguese", nativeName: "Português" },
  fr: { englishName: "French", nativeName: "Français" },
  de: { englishName: "German", nativeName: "Deutsch" },
  it: { englishName: "Italian", nativeName: "Italiano" },
  ja: { englishName: "Japanese", nativeName: "日本語" },
  ko: { englishName: "Korean", nativeName: "한국어" },
  "zh-CN": { englishName: "Chinese Simplified", nativeName: "简体中文" },
  "zh-TW": { englishName: "Chinese Traditional", nativeName: "繁體中文" },
  ru: { englishName: "Russian", nativeName: "Русский" },
  ar: { englishName: "Arabic", nativeName: "العربية" },
  hi: { englishName: "Hindi", nativeName: "हिन्दी" },
  th: { englishName: "Thai", nativeName: "ไทย" },
  vi: { englishName: "Vietnamese", nativeName: "Tiếng Việt" },
  tr: { englishName: "Turkish", nativeName: "Türkçe" },
  pl: { englishName: "Polish", nativeName: "Polski" },
  nl: { englishName: "Dutch", nativeName: "Nederlands" },
  sv: { englishName: "Swedish", nativeName: "Svenska" },
  no: { englishName: "Norwegian", nativeName: "Norsk" },
  da: { englishName: "Danish", nativeName: "Dansk" },
  fi: { englishName: "Finnish", nativeName: "Suomi" },
  hu: { englishName: "Hungarian", nativeName: "Magyar" },
  cs: { englishName: "Czech", nativeName: "Čeština" },
  sk: { englishName: "Slovak", nativeName: "Slovenčina" },
  ro: { englishName: "Romanian", nativeName: "Română" },
  bg: { englishName: "Bulgarian", nativeName: "Български" },
  hr: { englishName: "Croatian", nativeName: "Hrvatski" },
  sr: { englishName: "Serbian", nativeName: "Српски" },
  sl: { englishName: "Slovenian", nativeName: "Slovenščina" },
  et: { englishName: "Estonian", nativeName: "Eesti" },
  lv: { englishName: "Latvian", nativeName: "Latviešu" },
  lt: { englishName: "Lithuanian", nativeName: "Lietuvių" },
  el: { englishName: "Greek", nativeName: "Ελληνικά" },
  he: { englishName: "Hebrew", nativeName: "עברית" },
  fa: { englishName: "Persian", nativeName: "فارسی" },
  ur: { englishName: "Urdu", nativeName: "اردو" },
  bn: { englishName: "Bengali", nativeName: "বাংলা" },
  ta: { englishName: "Tamil", nativeName: "தமிழ்" },
  te: { englishName: "Telugu", nativeName: "తెలుగు" },
  kn: { englishName: "Kannada", nativeName: "ಕನ್ನಡ" },
  ml: { englishName: "Malayalam", nativeName: "മലയാളം" },
  gu: { englishName: "Gujarati", nativeName: "ગુજરાતી" },
  mr: { englishName: "Marathi", nativeName: "मराठी" },
  pa: { englishName: "Punjabi", nativeName: "ਪੰਜਾਬੀ" },
  ne: { englishName: "Nepali", nativeName: "नेपाली" },
  si: { englishName: "Sinhala", nativeName: "සිංහල" },
  my: { englishName: "Myanmar", nativeName: "မြန်မာ" },
  km: { englishName: "Khmer", nativeName: "ខ្មែរ" },
  lo: { englishName: "Lao", nativeName: "ລາວ" },
  ka: { englishName: "Georgian", nativeName: "ქართული" },
  am: { englishName: "Amharic", nativeName: "አማርኛ" },
  sw: { englishName: "Swahili", nativeName: "Kiswahili" },
  zu: { englishName: "Zulu", nativeName: "isiZulu" },
  af: { englishName: "Afrikaans", nativeName: "Afrikaans" },
  is: { englishName: "Icelandic", nativeName: "Íslenska" },
  mt: { englishName: "Maltese", nativeName: "Malti" },
  cy: { englishName: "Welsh", nativeName: "Cymraeg" },
  ga: { englishName: "Irish", nativeName: "Gaeilge" },
  eu: { englishName: "Basque", nativeName: "Euskara" },
  ca: { englishName: "Catalan", nativeName: "Català" },
  gl: { englishName: "Galician", nativeName: "Galego" },
  ast: { englishName: "Asturian", nativeName: "Asturianu" },
  br: { englishName: "Breton", nativeName: "Brezhoneg" },
  co: { englishName: "Corsican", nativeName: "Corsu" },
  eo: { englishName: "Esperanto", nativeName: "Esperanto" },
  la: { englishName: "Latin", nativeName: "Latina" },
  jv: { englishName: "Javanese", nativeName: "Basa Jawa" },
  su: { englishName: "Sundanese", nativeName: "Basa Sunda" },
  tl: { englishName: "Filipino", nativeName: "Filipino" },
  ceb: { englishName: "Cebuano", nativeName: "Cebuano" },
  haw: { englishName: "Hawaiian", nativeName: "ʻŌlelo Hawaiʻi" },
  mg: { englishName: "Malagasy", nativeName: "Malagasy" },
  sm: { englishName: "Samoan", nativeName: "Gagana Samoa" },
  to: { englishName: "Tongan", nativeName: "Lea Faka-Tonga" },
  fj: { englishName: "Fijian", nativeName: "Na Vosa Vakaviti" },
  mi: { englishName: "Maori", nativeName: "Te Reo Māori" },
  sn: { englishName: "Shona", nativeName: "chiShona" },
  st: { englishName: "Sotho", nativeName: "Sesotho" },
  xh: { englishName: "Xhosa", nativeName: "isiXhosa" },
  yo: { englishName: "Yoruba", nativeName: "Yorùbá" },
  ig: { englishName: "Igbo", nativeName: "Igbo" },
  ha: { englishName: "Hausa", nativeName: "Hausa" },
  rw: { englishName: "Kinyarwanda", nativeName: "Ikinyarwanda" },
  ny: { englishName: "Chichewa", nativeName: "Chichewa" },
  so: { englishName: "Somali", nativeName: "Soomaali" },
  ti: { englishName: "Tigrinya", nativeName: "ትግርኛ" },
  om: { englishName: "Oromo", nativeName: "Oromoo" },
  ak: { englishName: "Akan", nativeName: "Akan" },
  ee: { englishName: "Ewe", nativeName: "Eʋegbe" },
  tw: { englishName: "Twi", nativeName: "Twi" },
  lg: { englishName: "Luganda", nativeName: "Luganda" },
  ln: { englishName: "Lingala", nativeName: "Lingála" },
  kg: { englishName: "Kongo", nativeName: "Kikongo" },
  rn: { englishName: "Rundi", nativeName: "Ikirundi" },
  wo: { englishName: "Wolof", nativeName: "Wolof" },
  ff: { englishName: "Fulah", nativeName: "Fulfulde" },
  bm: { englishName: "Bambara", nativeName: "Bamanankan" },
  dyu: { englishName: "Dyula", nativeName: "Dyula" },
  kri: { englishName: "Krio", nativeName: "Krio" },
  luo: { englishName: "Luo", nativeName: "Dholuo" },
  gom: { englishName: "Goan Konkani", nativeName: "गोंयची कोंकणी" },
  sa: { englishName: "Sanskrit", nativeName: "संस्कृतम्" },
  pi: { englishName: "Pali", nativeName: "पालि" },
  bo: { englishName: "Tibetan", nativeName: "བོད་ཡིག" },
  dz: { englishName: "Dzongkha", nativeName: "རྫོང་ཁ" },
  ug: { englishName: "Uyghur", nativeName: "ئۇيغۇرچە" },
  kk: { englishName: "Kazakh", nativeName: "Қазақша" },
  ky: { englishName: "Kyrgyz", nativeName: "Кыргызча" },
  uz: { englishName: "Uzbek", nativeName: "Oʻzbekcha" },
  tk: { englishName: "Turkmen", nativeName: "Türkmençe" },
  tg: { englishName: "Tajik", nativeName: "Тоҷикӣ" },
  mn: { englishName: "Mongolian", nativeName: "Монгол" },
  ii: { englishName: "Sichuan Yi", nativeName: "ꆈꌠꉙ" },
  iu: { englishName: "Inuktitut", nativeName: "ᐃᓄᒃᑎᑐᑦ" },
  ik: { englishName: "Inupiaq", nativeName: "Iñupiatun" },
  chr: { englishName: "Cherokee", nativeName: "ᏣᎳᎩ" },
  chy: { englishName: "Cheyenne", nativeName: "Tsėhesenėstsestȯtse" },
  dak: { englishName: "Dakota", nativeName: "Dakȟótiyapi" },
  lkt: { englishName: "Lakota", nativeName: "Lakȟótiyapi" },
  nv: { englishName: "Navajo", nativeName: "Diné Bizaad" },
  qu: { englishName: "Quechua", nativeName: "Runa Simi" },
  gn: { englishName: "Guarani", nativeName: "Avañe'ẽ" },
  ay: { englishName: "Aymara", nativeName: "Aymar Aru" },
};

/**
 * Legacy fallback for backward compatibility
 * @deprecated Use LANGUAGES_FALLBACK_ENHANCED instead
 */
const LANGUAGES_FALLBACK: LanguageMap = Object.entries(
  LANGUAGES_FALLBACK_ENHANCED,
).reduce<LanguageMap>((legacy, [code, data]) => {
  legacy[code] = data.englishName;
  return legacy;
}, {});
