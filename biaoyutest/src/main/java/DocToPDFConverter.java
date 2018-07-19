import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.net.URL;


import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.convert.in.Doc;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.RFonts;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;


public class DocToPDFConverter extends Converter {


    public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
                             boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }


    @Override
    public void convert() throws Exception {
        loading();


        InputStream iStream = inStream;
        try {
            WordprocessingMLPackage wordMLPackage = getMLPackage(iStream);
            Mapper fontMapper = new IdentityPlusMapper();
            String fontFamily = "SimSun";

            /*Resource fileRource = new ClassPathResource("F:\\ymwork\\biaoyutest\\src\\main\\java\\simsun.ttc");
            String path =  fileRource.getFile().getAbsolutePath();*/
            URL fontUrl = new URL("file:"+"F:\\ymwork\\biaoyutest\\src\\main\\java\\simsun.ttc");
            PhysicalFonts.addPhysicalFont(fontUrl);

            PhysicalFont simsunFont = PhysicalFonts.get(fontFamily);
            fontMapper.put(fontFamily, simsunFont);


            RFonts rfonts = Context.getWmlObjectFactory().createRFonts(); // 设置文件默认字体
            rfonts.setAsciiTheme(null);
            rfonts.setAscii(fontFamily);
            wordMLPackage.getMainDocumentPart().getPropertyResolver().getDocumentDefaultRPr().setRFonts(rfonts);
            wordMLPackage.setFontMapper(fontMapper);
            FOSettings foSettings = Docx4J.createFOSettings();
            foSettings.setWmlPackage(wordMLPackage);
            Docx4J.toFO(foSettings, outStream, Docx4J.FLAG_EXPORT_PREFER_XSL);


        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            IOUtils.closeQuietly(outStream);
        }
  
  
/* 
* InputStream iStream = inStream; 
*  
*  
*  
* String regex = null; //Windows: // String 
* regex=".*(calibri|camb|cour|arial|symb|times|Times|zapf).*"; regex= 
* ".*(calibri|camb|cour|arial|times|comic|georgia|impact|LSANS|pala|tahoma|trebuc|verdana|symbol|webdings|wingding).*"; 
* // Mac // String // 
* regex=".*(Courier New|Arial|Times New Roman|Comic Sans|Georgia|Impact|Lucida Console|Lucida Sans Unicode|Palatino Linotype|Tahoma|Trebuchet|Verdana|Symbol|Webdings|Wingdings|MS Sans Serif|MS Serif).*" 
* ; PhysicalFonts.setRegex(regex); WordprocessingMLPackage 
* wordMLPackage = getMLPackage(iStream); // WordprocessingMLPackage 
* wordMLPackage = WordprocessingMLPackage.load(iStream) FieldUpdater 
* updater = new FieldUpdater(wordMLPackage); updater.update(true); // 
* process processing(); // Add font 
*  
* Mapper fontMapper = new IdentityPlusMapper(); 
*  
* PhysicalFont font = PhysicalFonts.get("Arial UTF-8 MS"); if (font != 
* null) { fontMapper.put("Times New Roman", font); 
* fontMapper.put("Arial", font); fontMapper.put("Calibri", font); } 
* fontMapper.put("Calibri", PhysicalFonts.get("Calibri")); 
* fontMapper.put("Algerian", font); fontMapper.put("华文行楷", 
* PhysicalFonts.get("STXingkai")); fontMapper.put("华文仿宋", 
* PhysicalFonts.get("STFangsong")); fontMapper.put("隶书", 
* PhysicalFonts.get("LiSu")); fontMapper.put("Libian SC Regular", 
* PhysicalFonts.get("SimSun")); 
* wordMLPackage.setFontMapper(fontMapper); FOSettings foSettings = 
* Docx4J.createFOSettings(); foSettings.setFoDumpFile(new 
* java.io.File("E:/xi.fo")); foSettings.setWmlPackage(wordMLPackage); 
* // Docx4J.toPDF(wordMLPackage, outStream); Docx4J.toFO(foSettings, 
* outStream, Docx4J.FLAG_EXPORT_PREFER_XSL); 
*/
        finished();


    }


    protected WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception {
//PrintStream originalStdout = System.out;  


        System.setOut(new PrintStream(new OutputStream() {
            public void write(int b) {
// DO NOTHING  
            }
        }));


        WordprocessingMLPackage mlPackage = Doc.convert(iStream);
//System.setOut(originalStdout);  
//System.out.println(outStream);  
        return mlPackage;
    }


}  