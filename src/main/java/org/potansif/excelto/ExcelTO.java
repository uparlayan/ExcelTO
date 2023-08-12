/**
 * ExcelTO, Excel dosyalarını JSON veya CSV formatına dönüştüren bir konsol uygulamasıdır.
 * Dönüşüm işlemleri için Apache POI ve JSON kütüphaneleri kullanılmıştır.
 * XML desteği? O kadar da gerekli değil...
 */
package org.potansif.excelto;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.text.Normalizer;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
  F:\Test\Java>ExcelTO.exe "F:\Test\Java\TEST.xls" "F:\Test\Java\TEST.json" 0 "0,1,2,3,4,5,6"
  F:\Test\Java>ExcelTO.exe "F:\Test\Java\TEST.xls" "F:\Test\Java\TEST.csv"  0 "0,1,2,3,4,5,6"
*/

/**
 * Bu sınıf, Excel dosyalarını CSV veya JSON formatına dönüştüren işlemleri gerçekleştirir.
 * @author Uğur Parlayan
 * @company Potansif Yazılım Hizmetleri
 * @since 09.08.2023
 * @version 0.1
 */
public class ExcelTO {

    // Karakter dönüşüm haritası ve diğer değişkenler burada yer alır...
    private static Map<String, String> characterMap = new HashMap<>();

    // Latih 1 ve Latin 1 Extended karakterlerini içerir.
    static {
        characterMap.put("À", "A");
        characterMap.put("Á", "A");
        characterMap.put("Â", "A");
        characterMap.put("Ã", "A");
        characterMap.put("Ä", "A");
        characterMap.put("Å", "A");
        characterMap.put("Æ", "AE");
        characterMap.put("Ç", "C");
        characterMap.put("È", "E");
        characterMap.put("É", "E");
        characterMap.put("Ê", "E");
        characterMap.put("Ë", "E");
        characterMap.put("Ì", "I");
        characterMap.put("Í", "I");
        characterMap.put("Î", "I");
        characterMap.put("Ï", "I");
        characterMap.put("Ð", "D");
        characterMap.put("Ñ", "N");
        characterMap.put("Ò", "O");
        characterMap.put("Ó", "O");
        characterMap.put("Ô", "O");
        characterMap.put("Õ", "O");
        characterMap.put("Ö", "O");
        characterMap.put("×", "x");
        characterMap.put("Ø", "O");
        characterMap.put("Ù", "U");
        characterMap.put("Ú", "U");
        characterMap.put("Û", "U");
        characterMap.put("Ü", "U");
        characterMap.put("Ý", "Y");
        characterMap.put("Þ", "Th");
        characterMap.put("ß", "ss");
        characterMap.put("à", "a");
        characterMap.put("á", "a");
        characterMap.put("â", "a");
        characterMap.put("ã", "a");
        characterMap.put("ä", "a");
        characterMap.put("å", "a");
        characterMap.put("æ", "ae");
        characterMap.put("ç", "c");
        characterMap.put("è", "e");
        characterMap.put("é", "e");
        characterMap.put("ê", "e");
        characterMap.put("ë", "e");
        characterMap.put("ì", "i");
        characterMap.put("í", "i");
        characterMap.put("î", "i");
        characterMap.put("ï", "i");
        characterMap.put("ð", "d");
        characterMap.put("ñ", "n");
        characterMap.put("ò", "o");
        characterMap.put("ó", "o");
        characterMap.put("ô", "o");
        characterMap.put("õ", "o");
        characterMap.put("ö", "o");
        characterMap.put("÷", "÷");
        characterMap.put("ø", "o");
        characterMap.put("ù", "u");
        characterMap.put("ú", "u");
        characterMap.put("û", "u");
        characterMap.put("ü", "u");
        characterMap.put("ý", "y");
        characterMap.put("þ", "th");
        characterMap.put("ÿ", "y");
        characterMap.put("Ā", "A");
        characterMap.put("ā", "a");
        characterMap.put("Ă", "A");
        characterMap.put("ă", "a");
        characterMap.put("Ą", "A");
        characterMap.put("ą", "a");
        characterMap.put("Ć", "C");
        characterMap.put("ć", "c");
        characterMap.put("Ĉ", "C");
        characterMap.put("ĉ", "c");
        characterMap.put("Ċ", "C");
        characterMap.put("ċ", "c");
        characterMap.put("Č", "C");
        characterMap.put("č", "c");
        characterMap.put("Ď", "D");
        characterMap.put("ď", "d");
        characterMap.put("Đ", "D");
        characterMap.put("đ", "d");
        characterMap.put("Ē", "E");
        characterMap.put("ē", "e");
        characterMap.put("Ĕ", "E");
        characterMap.put("ĕ", "e");
        characterMap.put("Ė", "E");
        characterMap.put("ė", "e");
        characterMap.put("Ę", "E");
        characterMap.put("ę", "e");
        characterMap.put("Ě", "E");
        characterMap.put("ě", "e");
        characterMap.put("Ĝ", "G");
        characterMap.put("ĝ", "g");
        characterMap.put("Ğ", "G");
        characterMap.put("ğ", "g");
        characterMap.put("Ġ", "G");
        characterMap.put("ġ", "g");
        characterMap.put("Ģ", "G");
        characterMap.put("ģ", "g");
        characterMap.put("Ĥ", "H");
        characterMap.put("ĥ", "h");
        characterMap.put("Ħ", "H");
        characterMap.put("ħ", "h");
        characterMap.put("Ĩ", "I");
        characterMap.put("ĩ", "i");
        characterMap.put("Ī", "I");
        characterMap.put("ī", "i");
        characterMap.put("Ĭ", "I");
        characterMap.put("ĭ", "i");
        characterMap.put("Į", "I");
        characterMap.put("į", "i");
        characterMap.put("İ", "I");
        characterMap.put("ı", "i");
        characterMap.put("Ĳ", "IJ");
        characterMap.put("ĳ", "ij");
        characterMap.put("Ĵ", "J");
        characterMap.put("ĵ", "j");
        characterMap.put("Ķ", "K");
        characterMap.put("ķ", "k");
        characterMap.put("ĸ", "k");
        characterMap.put("Ĺ", "L");
        characterMap.put("ĺ", "l");
        characterMap.put("Ļ", "L");
        characterMap.put("ļ", "l");
        characterMap.put("Ľ", "L");
        characterMap.put("ľ", "l");
        characterMap.put("Ŀ", "L");
        characterMap.put("ŀ", "l");
        characterMap.put("Ł", "L");
        characterMap.put("ł", "l");
        characterMap.put("Ń", "N");
        characterMap.put("ń", "n");
        characterMap.put("Ņ", "N");
        characterMap.put("ņ", "n");
        characterMap.put("Ň", "N");
        characterMap.put("ň", "n");
        characterMap.put("ŉ", "n");
        characterMap.put("Ŋ", "N");
        characterMap.put("ŋ", "n");
        characterMap.put("Ō", "O");
        characterMap.put("ō", "o");
        characterMap.put("Ŏ", "O");
        characterMap.put("ŏ", "o");
        characterMap.put("Ő", "O");
        characterMap.put("ő", "o");
        characterMap.put("Œ", "OE");
        characterMap.put("œ", "oe");
        characterMap.put("Ŕ", "R");
        characterMap.put("ŕ", "r");
        characterMap.put("Ŗ", "R");
        characterMap.put("ŗ", "r");
        characterMap.put("Ř", "R");
        characterMap.put("ř", "r");
        characterMap.put("Ś", "S");
        characterMap.put("ś", "s");
        characterMap.put("Ŝ", "S");
        characterMap.put("ŝ", "s");
        characterMap.put("Ş", "S");
        characterMap.put("ş", "s");
        characterMap.put("Š", "S");
        characterMap.put("š", "s");
        characterMap.put("Ţ", "T");
        characterMap.put("ţ", "t");
        characterMap.put("Ť", "T");
        characterMap.put("ť", "t");
        characterMap.put("Ŧ", "T");
        characterMap.put("ŧ", "t");
        characterMap.put("Ũ", "U");
        characterMap.put("ũ", "u");
        characterMap.put("Ū", "U");
        characterMap.put("ū", "u");
        characterMap.put("Ŭ", "U");
        characterMap.put("ŭ", "u");
        characterMap.put("Ů", "U");
        characterMap.put("ů", "u");
        characterMap.put("Ű", "U");
        characterMap.put("ű", "u");
        characterMap.put("Ų", "U");
        characterMap.put("ų", "u");
        characterMap.put("Ŵ", "W");
        characterMap.put("ŵ", "w");
        characterMap.put("Ŷ", "Y");
        characterMap.put("ŷ", "y");
        characterMap.put("Ÿ", "Y");
        characterMap.put("Ź", "Z");
        characterMap.put("ź", "z");
        characterMap.put("Ż", "Z");
        characterMap.put("ż", "z");
        characterMap.put("Ž", "Z");
        characterMap.put("ž", "z");
        characterMap.put("ſ", "s");
    }

    /**
     * Programın başlangıç noktası. Excel dosyasını JSON veya CSV formatına dönüştürür.
     *
     * @param args Komut satırı argümanları:
     *             args[0]: Excel dosyasının yolu
     *             args[1]: Çıkış dosyasının yolu
     *             args[2]: Sayfa indeksi
     *             args[3]: Dönüştürülecek kolon numaraları (virgülle ayrılmış)
     */
    public static void main(String[] args) {
        // Komut satırı argümanlarını kontrol ediyoruz
        if (args.length != 4) {
            System.out.println("Eksik parametre hatası");
            System.out.println("ExcelTO <Excel Dosyası> <Çıkış Dosyası> <Sheet indeksi> <kolon numaraları, virgüllü>");
            System.out.println("ExcelTO \"C:\\Alesta\\REQUEST_SABLON.xls\"  \"C:\\Test\\Java\\output.csv\"  0 \"0,1,2,3,4,5,6\"");
            System.out.println("ExcelTO \"C:\\Alesta\\REQUEST_SABLON.xlsx\" \"C:\\Test\\Java\\output.json\" 0 \"0,1,2,3,4,5,6\"");
            return;
        }
        String inputFilePath = args[0];
        String outputFilePath = args[1];
        int sheetIndex = Integer.parseInt(args[2]);
        String columnIndexes = args[3];

        String outputFormat = getOutputFormat(outputFilePath);

        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = getWorkbook(inputStream);
             FileOutputStream outputStream = new FileOutputStream(outputFilePath, false)) {

            // UTF-8 BOM karakterini yazıyoruz
            outputStream.write(0xEF);
            outputStream.write(0xBB);
            outputStream.write(0xBF);

            // Excel sayfasını alıyoruz
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            if (sheet == null) {
                System.out.println("Belirtilen sayfa indeksi (" + sheetIndex + ") Excel dosyasında bulunamadı.");
                return;
            }

            if ("CSV".equalsIgnoreCase(outputFormat)) {
                // CSV içeriğini oluşturuyoruz
                String csvContent = createCSVContent(sheet, columnIndexes);
                byte[] csvBytes = csvContent.getBytes(StandardCharsets.UTF_8); //("UTF-8");
                outputStream.write(csvBytes);
                System.out.println("Excel dosyası başarıyla CSV formatına dönüştürüldü. Çıkış dosyası yolu: " + outputFilePath);
            } else
            if ("JSON".equalsIgnoreCase(outputFormat)) {
                // JSON içeriğini oluşturuyoruz
                JSONArray jsonArray = createJSONArray(sheet, columnIndexes);
                byte[] jsonBytes = jsonArray.toString(2).getBytes(StandardCharsets.UTF_8); //("UTF-8");
                outputStream.write(jsonBytes);
                System.out.println("Excel dosyası başarıyla JSON formatına dönüştürüldü. Çıkış dosyası yolu: " + outputFilePath);
            } else {
                System.out.println("Geçersiz çıkış formatı. Lütfen 'CSV' veya 'JSON' kullanın.");
            }

        } catch (IOException | JSONException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * Verilen FileInputStream üzerinden bir Workbook nesnesi oluşturur.
     *
     * @param inputStream Excel dosyasını temsil eden FileInputStream
     * @return Workbook nesnesi
     * @throws IOException Eğer Workbook oluşturulurken bir hata oluşursa
     */
    private static Workbook getWorkbook(FileInputStream inputStream) throws IOException {
        return WorkbookFactory.create(inputStream);
    }

    /**
     * Verilen Sheet ve kolon indeksleri üzerinden bir CSV içeriği oluşturur.
     *
     * @param sheet         Excel sayfası
     * @param columnIndexes Virgülle ayrılmış kolon indeksleri
     * @return Oluşturulan CSV içeriği
     */
    private static String createCSVContent(Sheet sheet, String columnIndexes) {
        DataFormatter dataFormatter = new DataFormatter();
        StringBuilder csvContent = new StringBuilder();

        for (Row row : sheet) {
            String[] columnIndexArray = columnIndexes.replaceAll("\"", "").split(",");
            for (String columnIndexStr : columnIndexArray) {
                int columnIndex = Integer.parseInt(columnIndexStr);
                Cell cell = row.getCell(columnIndex);
                String cellValue = (cell == null) ? "" : dataFormatter.formatCellValue(cell);
                cellValue = convertToTurkishEquivalent(cellValue); // Türkçe karakter dönüşümü
                csvContent.append("\"").append(cellValue.trim()).append("\";");
            }
            csvContent.deleteCharAt(csvContent.length() - 1).append(System.lineSeparator());
        }

        return csvContent.toString();
    }

    /**
     * Türkçe karakterleri düzeltmek için kullanılan yardımcı fonksiyon.
     *
     * @param input Düzeltilecek girdi metni
     * @return Düzeltildikten sonra elde edilen metin
     */
    private static String convertToTurkishEquivalent(String input) {
        String normalizedInput = Normalizer.normalize(input, Normalizer.Form.NFD);
        StringBuilder convertedInput = new StringBuilder();

        for (int i = 0; i < normalizedInput.length(); i++) {
            String currentChar = String.valueOf(normalizedInput.charAt(i));
            String replacement = characterMap.getOrDefault(currentChar, currentChar);
            convertedInput.append(replacement);
        }

        return convertedInput.toString();
    }

    /**
     * Verilen Sheet ve kolon indeksleri üzerinden bir JSON dizisi oluşturur.
     *
     * @param sheet         Excel sayfası
     * @param columnIndexes Virgülle ayrılmış kolon indeksleri
     * @return Oluşturulan JSON dizisi
     * @throws JSONException JSON dizisi oluşturulurken bir hata oluşursa
     */
    private static JSONArray createJSONArray(Sheet sheet, String columnIndexes) throws JSONException {
        DataFormatter dataFormatter = new DataFormatter();
        JSONArray jsonArray = new JSONArray();

        for (Row row : sheet) {
            String[] columnIndexArray = columnIndexes.replaceAll("\"", "").split(",");
            JSONObject jsonObject = new JSONObject();
            for (String columnIndexStr : columnIndexArray) {
                int columnIndex = Integer.parseInt(columnIndexStr);
                Cell cell = row.getCell(columnIndex);
                String cellValue = (cell == null) ? "" : dataFormatter.formatCellValue(cell);
                cellValue = convertToTurkishEquivalent(cellValue); // Türkçe karakter dönüşümü
                jsonObject.put("Kolon " + columnIndex, cellValue.trim());
            }
            jsonArray.put(jsonObject);
        }

        return jsonArray;
    }

    /**
     * Verilen dosya yolundan çıkış formatını belirler.
     *
     * @param outputFilePath Çıkış dosyasının yolu
     * @return Dosya uzantısına göre belirlenen çıkış formatı
     */
    private static String getOutputFormat(String outputFilePath) {
        int dotIndex = outputFilePath.lastIndexOf(".");
        if (dotIndex > 0) {
            return outputFilePath.substring(dotIndex + 1).toUpperCase();
        }
        return "";
    }
}
