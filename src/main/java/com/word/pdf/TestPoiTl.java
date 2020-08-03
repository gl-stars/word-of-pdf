package com.word.pdf;

import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.util.HashMap;

/**
 * 基于 poi-tl 相关练习
 * @author: stars
 * @date: 2020年 08月 03日 11:49
 **/
public class TestPoiTl {

    public static void main(String[] args) throws IOException {

        fillWord();

    }
    public static void fillWord() throws IOException {
        //核心API采用了极简设计，只需要一行代码
        XWPFTemplate.compile("C:\\Users\\stars\\Desktop\\template.docx").render(new HashMap<String, Object>(){{
            put("123156465464", "贵州驰联科技有限公司哈哈");
            put("user.name", "贵州驰联科技有限公司哈哈");
        }}).writeToFile("C:\\Users\\stars\\Desktop\\out_template.docx");
    }
}
