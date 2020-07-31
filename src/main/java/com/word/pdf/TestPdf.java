package com.word.pdf;

import com.aspose.pdf.*;

import java.io.IOException;

/**
 * @author: stars
 * @date: 2020年 07月 30日 18:16
 **/
public class TestPdf {
    public static void main(String[] args) {
//        GetPageCountWithoutSavingPDF(7);
//        savingToDoc();
//        savingToDOCX();
        ConvertPDFToPPTX();
    }


    /**
     * 获取PDF指定页面
     * <p>如果获取所有PDF，就直接使用 pdfDocument。如果获取指定页数就是使用 newDocument。如果返回所有pdfDocument.save(response.getOutputStream());
     * 如果返回指定页，newDocument.save(response.getOutputStream());
     * </p>
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
    public static void savingToDoc(){
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
    public static void ConvertPDFToPPTX(){
        // Load PDF document
        Document doc =new Document("C:\\Users\\stars\\Desktop\\t.pdf");
        // Instantiate PptxSaveOptions instance
        PptxSaveOptions pptx_save = new PptxSaveOptions();
        // Save the output in PPTX format
        doc.save("C:\\Users\\stars\\Desktop\\output.pptx", pptx_save);
    }

}
