package com.word.pdf;

import com.aspose.pdf.*;
import com.aspose.pdf.text.CustomFontSubstitutionBase;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author: stars
 * @date: 2020年 07月 30日 18:16
 **/
public class TestPdf {
    public static void main(String[] args) throws Exception {
//        GetPageCountWithoutSavingPDF(7);
//        savingToDoc();
//        savingToDOCX();
//        ConvertPDFToPPTX();
//        ConvertPDFToSVGFormat();
//        ConvertSVGFileToPDFFormat();
//        DefaultFontWhenSpecificFontMissing();
        EscapeHTMLTagsAndSpecialCharacters();
    }


    /**
     * 获取PDF指定页面
     * <p>如果获取所有PDF，就直接使用 pdfDocument。如果获取指定页数就是使用 newDocument。如果返回所有pdfDocument.save(response.getOutputStream());
     * 如果返回指定页，newDocument.save(response.getOutputStream());
     * </p>
     *
     * @param page 获取的页数
     * @return pdf总页数
     */
    public static Integer GetPageCountWithoutSavingPDF(Integer page) throws IOException {
        // Open a document
        Document pdfDocument = new Document("C:\\Users\\stars\\Desktop\\a.pdf");
        // 获取指定页面
        Page pdfPage = pdfDocument.getPages().getUnrestricted(page);
        // 创建Document 对象，将指定页面写入
        Document newDocument = new Document();
        // Add the page to the Pages collection of new document object
        newDocument.getPages().add(pdfPage);
        // Save the new file
        newDocument.save("C:\\Users\\stars\\Desktop\\page_" + pdfPage.getNumber() + ".pdf");
        // 返回总页数
        return pdfDocument.getPages().size();
    }

    /**
     * pdf转为doc
     */
    public static void savingToDoc() {
        // 读取pdf文档
        Document pdfDocument = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        // 转换为doc文档
        pdfDocument.save("C:\\Users\\stars\\Desktop\\TableHeightIssue.doc", SaveFormat.Doc);
    }

    /**
     * pdf转为docx
     */
    public static void savingToDOCX() {
        // Load source PDF file
        Document doc = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        // Instantiate Doc SaveOptions instance
        DocSaveOptions saveOptions = new DocSaveOptions();
        // Set output file format as DOCX
        saveOptions.setFormat(DocSaveOptions.DocFormat.DocX);
        // Save resultant DOCX file
        doc.save("C:\\Users\\stars\\Desktop\\TableHeightIssue.docx", saveOptions);
    }


    /**
     * 将PDF转为ppt
     */
    public static void ConvertPDFToPPTX() {
        // Load PDF document
        Document doc = new Document("C:\\Users\\stars\\Desktop\\t.pdf");
        // Instantiate PptxSaveOptions instance
        PptxSaveOptions pptx_save = new PptxSaveOptions();
        // Save the output in PPTX format
        doc.save("C:\\Users\\stars\\Desktop\\output.pptx", pptx_save);
    }

    /**
     * 将pdf转为SVG图片
     */
    public static void ConvertPDFToSVGFormat() {
        // load PDF document
        Document doc = new Document("C:\\Users\\stars\\Desktop\\t.pdf");
        // instantiate an object of SvgSaveOptions
        SvgSaveOptions saveOptions = new SvgSaveOptions();
        // do not compress SVG image to Zip archive
        saveOptions.CompressOutputToZipArchive = false;
        // resultant file name
        String outFileName = "C:\\Users\\stars\\Desktop\\Output.svg";
        // save the output in SVG files
        doc.save(outFileName, saveOptions);
    }

    /**
     * 将SVG转为PDF
     */
    public static void ConvertSVGFileToPDFFormat() {
        String file = "C:\\Users\\stars\\Desktop\\Output.svg";
        // Instantiate LoadOption object using SVG load option
        LoadOptions options = new SvgLoadOptions();
        // Create Document object
        Document document = new Document(file, options);
        // Save the resultant PDF document
        document.save("C:\\Users\\stars\\Desktop\\Result.pdf");
    }

    /**
     * pdf转为HTML
     * @throws Exception
     */
    public static void DefaultFontWhenSpecificFontMissing() throws Exception {
        String myDir = "C:\\Users\\stars\\Desktop\\";
        Document pdf = new Document(myDir + "t.pdf");
        // configure font substitution
        CustomSubst1 subst1 = new CustomSubst1();
        FontRepository.getSubstitutions().add(subst1);
        // Configure notifier to console
        pdf.FontSubstitution.add(new Document.FontSubstitutionHandler() {
            public void invoke(Font font, Font newFont) {
                // print substituted FontNames into console
                System.out.println("Warning: Font " + font.getFontName() + " was substituted with another font -> " + newFont.getFontName());
            }
        });
        HtmlSaveOptions htmlSaveOps = new HtmlSaveOptions();
        pdf.save(myDir + "Redis_1150_substitutedWithMSGothic_release.html", htmlSaveOps);
    }
    private static class CustomSubst1 extends CustomFontSubstitutionBase {
        public boolean trySubstitute(OriginalFontSpecification originalFontSpecification, /* out */com.aspose.pdf.Font[] substitutionFont) {
            substitutionFont[0] = FontRepository.findFont("MSGothic");
            return true;
        }
    }

    /**
     * pdf转为HTML格式
     */
    public static void EscapeHTMLTagsAndSpecialCharacters(){
        // Load existing PDf file
        Document pdfDoc = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        final Map names = new HashMap();
        /*pdfDoc.FontSubstitution.add(new Document.FontSubstitutionHandler() {
            public void invoke(Font font, Font newFont) {
                // add substituted FontNames into map.
                names.put(font.getFontName(), newFont.getFontName());
                // or print the message into console
                System.out.println("Warning: Font " + font.getFontName() + " was substituted with another font -> " + newFont.getFontName());
            }
        });*/
        // instantiate HTMLSave option to save output in HTML
        HtmlSaveOptions htmlSaveOps = new HtmlSaveOptions();
        // save resultant file
        pdfDoc.save("C:\\Users\\stars\\Desktop\\output.html", htmlSaveOps);
    }

}
