package com.word.pdf;

import com.aspose.words.*;

import java.util.regex.Pattern;

/**
 * @author: stars
 * @date: 2020年 07月 30日 13:58
 **/
public class TestDemo {

    public static void main(String[] args) throws Exception {
        String inputPath = "C:\\Users\\stars\\Desktop\\powerdesigner使用.doc" ;
        String outPath = "C:\\Users\\stars\\Desktop";
        String fileName = "powerdesigner使用";
//        setShowInBalloons(inputPath,outPath,fileName);
//        AppendDocuments();
        FindAndReplace("namea","这是标题");
        FindAndReplace("name","替换后的数据");
    }
    /**
     * 将word文档转为PDF文档
     *
     * @param inputPath word文档所在的地址
     * @param outPath 生成pdf输出的地址
     * @param fileName 生成后PDF的名字<b>注意：名字不要带后缀，下面自动封装了。</b>
     * @throws Exception
     */
    private static void setShowInBalloons(String inputPath,String outPath,String fileName) throws Exception {
        Document doc = new Document(inputPath);
        // 获取控制修订外观的RevisionOptions对象
        RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();
        // 在引出序号中显示删除修订
        revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
        // 输出
        doc.save(outPath + "\\" + fileName+ ".pdf");
    }

    /**
     * word文件追加
     */
    public static void AppendDocuments() throws Exception{
        // 文件路径
        String dataDir = "C:\\Users\\stars\\Desktop\\" ;
        // 源文件
        Document dstDoc = new Document(dataDir + "powerdesigner使用.doc");
        // 需要追加的文件（将srcDoc文件追加在dstDoc后面）
        Document srcDoc = new Document(dataDir + "ZS-1000质量手册.docx");
        // Append the source document to the destination document while keeping the original formatting of the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        // 追加后的文件
        dstDoc.save(dataDir + "TestFile Out.docx");
        //ExEnd:

        System.out.println("Documents appended successfully.");
    }

    /**
     * 查找并替换
     * @param replaceMark 替换表示符
     * @param replaceAfter 替换的数据
     * @throws Exception
     */
    public static void FindAndReplace(String replaceMark,String replaceAfter) throws Exception {
        // 文件路径
        String dataDir = "C:\\Users\\stars\\Desktop\\" ;
        // 源文件
        Document doc = new Document(dataDir + "ZS-1000质量手册.docx");
        // 检查文档的文本
        System.out.println("Original document text: " + doc.getRange().getText());

        Pattern regex = Pattern.compile(replaceMark, Pattern.CASE_INSENSITIVE);
        // Replace the text in the document.
        doc.getRange().replace(regex, replaceAfter, new FindReplaceOptions());
        // Check the replacement was made.
        System.out.println("Document text after replace: " + doc.getRange().getText());
        // 替换后的文件
        doc.save(dataDir + "ReplaceSimpleOut.doc");
    }
}
