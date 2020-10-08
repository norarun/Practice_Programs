/*Code*/
tesseract.setLanguage("jpn");
List<Word> wordList = tesseract.getWords(image, TessPageIteratorLevel.RIL_BLOCK);
String str = tesseract.doOCR(image);
        
/*Error message*/
ParamsModel::Incomplete line
