package com.word.pdf;

import com.aspose.words.*;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * @author: stars
 * @date: 2020年 07月 30日 13:58
 **/
public class TestDemo {

    public static void main(String[] args) throws Exception {
//        String inputPath = "C:\\Users\\stars\\Desktop\\ZS-1000质量手册.docx" ;
//        String outPath = "C:\\Users\\stars\\Desktop";
//        String fileName = "ZS-1000质量手册";
//        setShowInBalloons(inputPath,outPath,fileName);
//        AppendDocuments();
//        FindAndReplace("namea","这是标题");
//        FindAndReplace("name","替换后的数据");
//        embedCoreFonts();
        Map<String,String> map = new HashMap<>();
        map.put("1289114469835157857","贵州驰联科技有限公司");
        map.put("1289114469835157858","质 量 手 册");
        map.put("1289114469835157859","质量手册实施令");
        map.put("1289114469835157860","公正性和保密声明");
        map.put("1289114469835157862","质量方针目标和承诺");
        map.put("1289114469835157863","质量管理体系有效性运行承诺书");
        map.put("1289114469835157864","公司概况");
        map.put("1289114469835157866","质量管理组织机构图");
        map.put("1289114469835157865","质量管理职能分配表");
        map.put("1289114469835157868","GZCL/CX/2031-2020《检验检测报告管理程序》");
        map.put("1289114469835157869","修订的报告要注以唯一性标识。\n" +
                "4.5.11.3对本公司已经发签发的检测报告的实质性修改，公司由质量负责人追加报告更换的形式进行，并做如下声明：“对检测报告的补充，系列号…”，修改应满足本手册的所有要求。若有必要发布全新的检测报告时，应注以唯一性标识，并注明所替代的原件。");
        map.put("1289114469835157873","章 节 号");
        FindAndReplaceOfPdf(map);
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
    public static void findAndReplaceOfPdf(String replaceMark,String replaceAfter) throws Exception {
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

    /**
     * 更改字体
     * <p>将word文档转为PDF文档，设置字体</p>
     * @throws Exception
     */
    public static void embedCoreFonts() throws Exception {
        // 文件路径
        String dataDir = "C:\\Users\\stars\\Desktop\\" ;
        ///ExStart:embedCoreFonts
        // Load the document to render.
        Document doc = new Document(dataDir + "powerdesigner使用.doc");

        // To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
        PdfSaveOptions options = new PdfSaveOptions();
        // true 使用核心字体
        options.setUseCoreFonts(true);

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc.save(dataDir + "Rendering.DisableEmbedWindowsFonts_Out.pdf");
        //ExEnd:embedCoreFonts
    }

    /**
     * 查找并替换并转为pdf文档保存
     * @param maps
     * @throws Exception
     */
    public static void FindAndReplaceOfPdf(Map<String,String > maps) throws Exception {

        // 文件路径
        String dataDir = "C:\\Users\\stars\\Desktop\\";
        // 源文件
        Document doc = new Document(dataDir + "ZS-1000质量手册.docx");
        maps.forEach((k, v) -> {
            try {
                Pattern regex = Pattern.compile(k, Pattern.CASE_INSENSITIVE);
                doc.getRange().replace(regex, v, new FindReplaceOptions());
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
        // 获取控制修订外观的RevisionOptions对象
        RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();
        // 在引出序号中显示删除修订
        revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
        // 输出
        doc.save(dataDir + "\\admin.pdf");
    }
}
