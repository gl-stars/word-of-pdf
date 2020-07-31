package com.word.pdf.controller;

import com.aspose.pdf.Document;
import com.aspose.pdf.Page;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

/**
 * @author: stars
 * @date: 2020年 07月 30日 16:37
 **/
@RestController
public class PdfAndWordController {

    /**
     * 获取所有PDF
     * @param response
     * @param page
     * @throws Exception
     */
    @GetMapping("/pdf")
    public void getWordOfPdf(HttpServletResponse response,Integer page) throws Exception{
        // Open a document
        Document pdfDocument = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
       pdfDocument.save(response.getOutputStream());
    }

    /**
     * 获取指定页面的PDF
     * <b>这种方式有些时候会出现问题，例如：指定某页，其实这个页数有的，但是有些时候就是不能读取出来。</b>
     * @param response
     * @param page
     * @throws Exception
     */
    @GetMapping("{page}")
    public void getPagePdf(HttpServletResponse response,@PathVariable Integer page) throws Exception{
        // Open a document
        Document pdfDocument = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        // 获取指定页面
        Page pdfPage = pdfDocument.getPages().getUnrestricted(page);
        // 创建Document 对象，将指定页面写入
        Document newDocument = new Document();
        // Add the page to the Pages collection of new document object
        newDocument.getPages().add(pdfPage);
        newDocument.save(response.getOutputStream());
    }
}
