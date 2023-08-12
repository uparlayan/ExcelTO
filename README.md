# ExcelTO
`ExcelTO`, Excel dosyalarını `JSON` veya `CSV` formatına dönüştüren bir konsol uygulamasıdır. Paket yönetimi için `maven` kullanılan projede dönüşüm işlemleri için `Apache POI` ve `JSON` paketleri kullanılmıştır. 

## JRE
https://adoptium.net/temurin/releases/ adresinden ise uyumlu `JRE` sürümünü indirebilirsiniz. Projede kullanılan JRE sürümü `17 - LTS`'dir.

## Dağıtım
https://launch4j.sourceforge.net/ adresinden `Launch4J` 3.50 sürümünü kullanarak uygulamayı `.EXE` türüne dönüştürebilir ve dağıtımını kolayca yapabilirsiniz. Gerekli ayarlamalar `L4jConfig.xml` dosyasında yer almaktadır. 

## Kullanım Örneği (Windows)
```powershell
C:\> ExcelTO.exe "F:\Temp\INPUT.xls" "F:\Data\OUTPUT.json" 0 "0,1,2,3,4,5,6"
```
1. Parametre excel dosyasını ifade eder.
2. Dönüştürülecek olan dosyanın tam adını ifade eder. dosya uzantısı CSV veya JSON değil ise hata alırsınız.
3. Sheet (sayfa) numarasıdır. Sıfırdan başlar
4. Kolonların sıra numarasıdır. A kolonu için sıfır kullanılır, B için bir ve bu böyle devam eder.
