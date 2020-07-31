package com.word.pdf.controller;

import com.aspose.pdf.Document;
import com.aspose.pdf.Page;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.List;

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
     * <b>这种方式有些时候会出现问题，例如：指定某页，其实这个页数有的，但是有些时候就是不能读取出来。
     * 推荐使用{@link PdfAndWordController#splitPdf(javax.servlet.http.HttpServletResponse, int, java.lang.String)}</b>
     *
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

    @GetMapping("/pdf/{page}")
    public void findPdf(HttpServletResponse response, @PathVariable int page) throws IOException{
        splitPdf(response,page,"C:\\Users\\stars\\Desktop\\abc.pdf");
    }

    /**
     *
     * @param pageNum 读取的页数
     * @param filePath 文件路径
     * @return
     */
    public void splitPdf(HttpServletResponse response, int pageNum, String filePath) {
        // 这是对应文件名
        File indexFile = new File(filePath);
        PDDocument document = null;
        try {
            document = PDDocument.load(indexFile);
            // 获取总页数
            int numberOfPages = document.getNumberOfPages();
            System.out.println("总页数告诉你哦："+numberOfPages);

            Splitter splitter = new Splitter();
            // 开始页
            splitter.setStartPage(pageNum);
            // 结束页
            splitter.setEndPage(pageNum);
            List<PDDocument> pages = splitter.split(document);
            for (PDDocument pdDocument : pages){
                pdDocument.save(response.getOutputStream());
                pdDocument.close();
            }
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
