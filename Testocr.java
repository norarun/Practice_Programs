/**
 *
 * @author noraru
 */
package OCR;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.List;
import javax.imageio.ImageIO;
import net.sourceforge.tess4j.ITessAPI.TessPageIteratorLevel;
import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import net.sourceforge.tess4j.Word;

public class TestOcr {
 
    public static void main(String[] args) throws IOException, TesseractException {
 
        //read
        File target = new File("src\\OCR\\receiptTrimingTOP.jpg");
        BufferedImage image = ImageIO.read(target);
 
        //analysis
        ITesseract tesseract = new Tesseract();
        tesseract.setDatapath("C:\\tessdata");
        tesseract.setLanguage("jpn");           //language (now error)
        List<Word> wordList = tesseract.getWords(image, TessPageIteratorLevel.RIL_BLOCK);
        String str = tesseract.doOCR(image);
 
        //output
        System.out.println(wordList);
        System.out.println(str);
 
    }
}
