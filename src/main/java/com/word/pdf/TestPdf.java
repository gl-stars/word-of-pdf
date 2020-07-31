package com.word.pdf;

import com.aspose.pdf.*;
import com.aspose.pdf.devices.EmfDevice;
import com.aspose.pdf.devices.Resolution;
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
//        EscapeHTMLTagsAndSpecialCharacters();
//        PDFToEMF();
//        addTableInExistingPDFDocument();
//        setAutoFitToWindowPropertyInColumnAdjustmentTypeEnumeration();
//        ForceTableRenderingOnNewPage();
        replaceTextOnAllPages("A","C");
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
     * <b>这种方式转换效果不好，推荐使用{@link TestPdf#EscapeHTMLTagsAndSpecialCharacters()}</b>
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


    /**
     * pdf转为emf文件，只能转换一页
     */
    public static void PDFToEMF(){
        // instantiate EmfDevice object
        EmfDevice device = new EmfDevice(new Resolution(96));
        // load existing PDF file
        Document doc = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        // save first page of PDF file as Emf image
        device.process(doc.getPages().get_Item(1), "C:\\Users\\stars\\Desktop\\output.emf");
    }

    /**
     * 在现有的PDF 文档中添加数据
     * <p>指定页添加表格</p>
     */
//    public static void addTableInExistingPDFDocument(){
//        // Load source PDF document
//        Document doc =  new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
//        // 初始化表的新实例
//        Table table = new Table();
//        // 将表格边框颜色设置为浅灰色
//        table.setBorder(new BorderInfo(BorderSide.All, .5f, Color.getLightGray()));
//        // 设置表格单元格的边框
//        table.setDefaultCellBorder(new BorderInfo(BorderSide.All, .5f, Color.getLightGray()));
//
//        // 添加三列十行的数据
//        for (int row_count = 1; row_count < 10; row_count++) {
//            // 向表中添加行
//            Row row = table.getRows().add();
//            // 添加表格的行数据
//            row.getCells().add("Column (" + row_count + ", 1)");
//            row.getCells().add("Column (" + row_count + ", 2)");
//            row.getCells().add("Column (" + row_count + ", 3)");
//        }
//        // 在PDF文档的第二页添加 table 这个表格
//        doc.getPages().getUnrestricted(2).getParagraphs().add(table);
//        doc.save( "C:\\Users\\stars\\Desktop\\Annotation_output.pdf");
//    }

    /**
     * 在现有的PDF 文档中添加数据
     * <p>指定页添加表格</p>
     */
    public static void addTableInExistingPDFDocument(){
        // Load source PDF document
        Document doc =  new Document("C:\\Users\\stars\\Desktop\\cc.pdf");
        // 初始化表的新实例
        Table table = new Table();
        // 将表格边框颜色设置为浅灰色
        table.setBorder(new BorderInfo(BorderSide.All, .5f, Color.getLightGray()));
        // 设置表格单元格的边框
        table.setDefaultCellBorder(new BorderInfo(BorderSide.All, .5f, Color.getLightGray()));
        // 创建行对象
        Row row = table.getRows().add();
        // 这一行中第一列的数据
        row.getCells().add("编号");
        // 这一行中第二列的数据
        row.getCells().add("姓名");
        // 这一行中第三列的数据
        row.getCells().add("性别");
        // 这一行中第四列的数据
        row.getCells().add("家庭住址");
        for (int i = 0 ;i<10;i++){
            Row rowLine = table.getRows().add();
            rowLine.getCells().add("" + i);
            rowLine.getCells().add("小花花");
            rowLine.getCells().add("女");
            rowLine.getCells().add("这里的家庭地址在怎么不全面");
        }
        // 在PDF文档的第二页添加 table 这个表格
        doc.getPages().getUnrestricted(1).getParagraphs().add(table);
        doc.save( "C:\\Users\\stars\\Desktop\\Annotation_output.pdf");
    }

    /**
     * 创建PDF文档，并添加表格
     */
    public static void setAutoFitToWindowPropertyInColumnAdjustmentTypeEnumeration(){
        //Instantiate the PDF object by calling its empty constructor
        Document doc = new Document();
        //Create the section in the PDF object
        Page page = doc.getPages().add();

        //Instantiate a table object
        Table tab = new Table();
        //Add the table in paragraphs collection of the desired section
        page.getParagraphs().add(tab);

        //Set with column widths of the table
        tab.setColumnWidths("50 50 50");
        tab.setColumnAdjustment(ColumnAdjustment.AutoFitToWindow);

        //Set default cell border using BorderInfo object
        tab.setDefaultCellBorder(new com.aspose.pdf.BorderInfo(com.aspose.pdf.BorderSide.All, 0.1F));

        //Set table border using another customized BorderInfo object
        tab.setBorder(new com.aspose.pdf.BorderInfo(com.aspose.pdf.BorderSide.All, 1F));
        //Create MarginInfo object and set its left, bottom, right and top margins
        com.aspose.pdf.MarginInfo margin = new com.aspose.pdf.MarginInfo();
        margin.setTop(5f);
        margin.setLeft(5f);
        margin.setRight(5f);
        margin.setBottom(5f);

        //Set the default cell padding to the MarginInfo object
        tab.setDefaultCellPadding(margin);

        //Create rows in the table and then cells in the rows
        com.aspose.pdf.Row row1 = tab.getRows().add();
        row1.getCells().add("col1");
        row1.getCells().add("col2");
        row1.getCells().add("col3");
        com.aspose.pdf.Row row2 = tab.getRows().add();
        row2.getCells().add("item1");
        row2.getCells().add("item2");
        row2.getCells().add("item3");

        //Save the PDF
        doc.save( "C:\\Users\\stars\\Desktop\\Annotation_output.pdf");
    }

    /**
     * 新建PDF添加数据
     */
    public static void ForceTableRenderingOnNewPage(){
        // Added document
        Document doc = new Document();
        PageInfo pageInfo = doc.getPageInfo();
        MarginInfo marginInfo = pageInfo.getMargin();
        marginInfo.setLeft(37);
        marginInfo.setRight(37);
        marginInfo.setTop(37);
        marginInfo.setBottom(37);
        pageInfo.setLandscape(true);
        Table table = new Table();
        table.setColumnWidths("50 100");
        // Added page.
        Page curPage = doc.getPages().add();
        for (int i = 1; i <= 120; i++) {
            Row row = table.getRows().add();
            row.setFixedRowHeight(15);
            Cell cell1 = row.getCells().add();
            cell1.getParagraphs().add(new TextFragment("Content 1"));
            Cell cell2 = row.getCells().add();
            cell2.getParagraphs().add(new TextFragment("HHHHH"));
        }
        Paragraphs paragraphs = curPage.getParagraphs();
        paragraphs.add(table);
        /********************************************/
        Table table1 = new Table();
        table.setColumnWidths("100 100");
        for (int i = 1; i <= 10; i++) {
            Row row = table1.getRows().add();
            Cell cell1 = row.getCells().add();
            cell1.getParagraphs().add(new TextFragment("LAAAAAAA"));
            Cell cell2 = row.getCells().add();
            cell2.getParagraphs().add(new TextFragment("LAAGGGGGG"));
        }
        table1.setInNewPage(true);
        // I want to keep table 1 to next page please...
        paragraphs.add(table1);
        //Save the PDF
        doc.save( "C:\\Users\\stars\\Desktop\\Annotation_output.pdf");
    }

    /**
     * 替换文本
     * <b>如果一个文字替换一个文字，还行，因为其他格式是不变话的。只能替换前三页</b>
     * @param BeforeReplacement 替换前
     * @param AfterReplacement 替换后
     */
    public static void replaceTextOnAllPages(String BeforeReplacement,String AfterReplacement){
        // 需要替换的文本
        Document doc = new Document("C:\\Users\\stars\\Desktop\\abc.pdf");
        // 需要替换的字符
        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(BeforeReplacement);
        // Accept the absorber for first page of document
        doc.getPages().accept(textFragmentAbsorber);
        // Get the extracted text fragments into collection
        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.getTextFragments();
        // Loop through the fragments
        for (TextFragment textFragment : (Iterable<TextFragment>) textFragmentCollection) {
            // Update text and other properties
            textFragment.setText(AfterReplacement);
            // 字体大小
            textFragment.getTextState().setFontSize(10);
            textFragment.getTextState().setForegroundColor(Color.getBlack());
            // 背景色，现在是白色
            textFragment.getTextState().setBackgroundColor(Color.getWhite());
        }
        // 替换后的文本
        doc.save( "C:\\Users\\stars\\Desktop\\Annotation_output.pdf");
    }


}
