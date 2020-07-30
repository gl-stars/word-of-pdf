package com.word.pdf.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

/**
 * @author: stars
 * @date: 2020年 07月 30日 16:37
 **/
@RestController
public class PdfAndWordController {

    @GetMapping("/pdf")
    public void getWordOfPdf(HttpServletResponse response) throws Exception{
    }
}
