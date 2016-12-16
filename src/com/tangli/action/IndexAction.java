package com.tangli.action;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.common.BitMatrix;
import com.lowagie.text.*;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.rtf.RtfWriter2;
import com.opensymphony.xwork2.ActionSupport;
import com.tangli.util.MatrixToImageWriter;

/**
 * Created by butta-hp on 2016/12/15.
 */

import java.io.*;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.stream.FileImageInputStream;
import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.ServiceUI;
import javax.print.SimpleDoc;
import javax.print.attribute.DocAttributeSet;
import javax.print.attribute.HashDocAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.swing.JFileChooser;


public  class IndexAction extends ActionSupport {

        public String mya() throws IOException,DocumentException{

            //接受传过来的参数

            String name="唐俪";
            String tid="31000001000061005";
            String sex="女";
            String adrress="中国";
            String adate="2016年10月05日";
            String jadrress="上海";
            String bdate="2016年11月25日";
            String jaddress="上海市公益服务促进中心";


            /**
             * 1.得到tid将其转换为二维码存储到计算机硬盘中
             * 2.将传过来的参数转换为
             */

            try {

//                String content = "120605181003";

                MultiFormatWriter multiFormatWriter = new MultiFormatWriter();

                Map hints = new HashMap();
                hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");
                //将传过来的tid
                BitMatrix bitMatrix = multiFormatWriter.encode(tid, BarcodeFormat.QR_CODE, 80, 80,hints);
                File file1 = new File("d:/","myzxing.jpeg");
                MatrixToImageWriter.writeToFile(bitMatrix, "jpeg", file1);

            } catch (Exception e) {
                e.printStackTrace();
            }
            // 设置纸张大小
            Document document = new Document(PageSize.A4);

            // 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
            // ByteArrayOutputStream baos = new ByteArrayOutputStream();

            File myfile = new File("D:/report.doc");

            RtfWriter2.getInstance(document, new FileOutputStream(myfile));

            document.open();

            // 设置中文字体

            BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA,
                    BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

            // 标题字体风格

            Font titleFont = new Font(bfChinese, 12, Font.BOLD);

            // // 正文字体风格
            //
            Font contextFont = new Font(bfChinese, 10, Font.NORMAL);

//            Paragraph title = new Paragraph("统计报告");
            //
            // 设置标题格式对齐方式

//            title.setAlignment(Element.ALIGN_CENTER);

            // title.setFont(titleFont);

//            document.add(title);

            BaseFont bf = BaseFont.createFont( "STSong-Light",   "UniGB-UCS2-H",   false,   false,   null,   null);
            Font   fontChinese5  =   new   Font(bf,8);
            PdfPTable table1 = new PdfPTable(2); //表格两列
            table1.setHorizontalAlignment(Element.ALIGN_CENTER); //垂直居中
            table1.setWidthPercentage(67);//表格的宽度为100%
            float[] wid1 ={0.55f,0.45f}; //两列宽度的比例
            table1.setWidths(wid1);
            table1.getDefaultCell().setBorderWidth(0); //不显示边框

            StringBuffer stringBuffer1=new StringBuffer();

            stringBuffer1.append("\n\n\n"+name+"  ");
            stringBuffer1.append(tid);
            stringBuffer1.append("  ");

            Paragraph context1 = new Paragraph(stringBuffer1.toString());

            // 正文格式左对齐

            context1.setAlignment(Element.ALIGN_BASELINE);

            // context.setFont(contextFont);

            // 离上一段落（标题）空的行数

            context1.setSpacingBefore(5);

            // 设置第一行空的列数

            context1.setFirstLineIndent(20);

            table1.addCell(context1);

            String imagepath = "d:/myzxing.jpeg";
            Image image = Image.getInstance(imagepath);
            table1.addCell(image);
            /**
             * 第一行结束
             */


            PdfPTable table2 = new PdfPTable(2); //表格两列
            table2.setHorizontalAlignment(Element.ALIGN_CENTER); //垂直居中
            table2.setWidthPercentage(67);//表格的宽度为100%
            float[] wid2 ={0.60f,0.40f}; //两列宽度的比例
            table2.setWidths(wid2);
            table2.getDefaultCell().setBorderWidth(0); //不显示边框


            StringBuffer stringBuffer2=new StringBuffer();

            stringBuffer2.append(sex+"          ");
            stringBuffer2.append(adrress+"          ");
            stringBuffer2.append(" ");


            Paragraph context2 = new Paragraph(stringBuffer2.toString());

            // 正文格式左对齐

            context2.setAlignment(Element.ALIGN_LEFT);

            // 离上一段落（标题）空的行数
            context2.setSpacingBefore(5);

            // 设置第一行空的列数
            context2.setFirstLineIndent(20);
            table2.addCell(context2);
            table2.addCell(adate+"\n\n");

            Paragraph context3 = new Paragraph(jadrress);

            // 正文格式左对齐

            context3.setAlignment(Element.ALIGN_LEFT);

            // 离上一段落（标题）空的行数
            context3.setSpacingBefore(5);

            // 设置第一行空的列数
            context3.setFirstLineIndent(20);
            table2.addCell(context3);
            table2.addCell(bdate+"\n");



            PdfPTable table3 = new PdfPTable(1); //表格两列
            table3.setHorizontalAlignment(Element.ALIGN_CENTER); //垂直居中
            table3.setWidthPercentage(67);//表格的宽度为100%
            table3.getDefaultCell().setBorderWidth(0); //不显示边框
            Paragraph context4 = new Paragraph(jaddress);

            // 正文格式左对齐

            context4.setAlignment(Element.ALIGN_LEFT);

            // 离上一段落（标题）空的行数
            context4.setSpacingBefore(5);

            // 设置第一行空的列数
            context4.setFirstLineIndent(20);
            table3.addCell(context4);



            document.add(table1);
            document.add(table2);
            document.add(table3);

            document.close();



            /**
             * 打印
             */
            //得到文本内容并打印
            JFileChooser fileChooser = new JFileChooser(); // 创建打印作业
            File file = new File("d:/report.doc"); // 获取选择的文件
            // 构建打印请求属性集
            HashPrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
            // 设置打印格式，因为未确定类型，所以选择autosense
            DocFlavor flavor = DocFlavor.INPUT_STREAM.PCL;
            // 定位默认的打印服务
            PrintService defaultService = PrintServiceLookup.lookupDefaultPrintService();
            InputStream fis = null;
            try {
                DocPrintJob job = defaultService.createPrintJob(); // 创建打印作业
                fis = new FileInputStream(file); // 构造待打印的文件流
                DocAttributeSet das = new HashDocAttributeSet();
                Doc doc = new SimpleDoc(fis, flavor, das);
                job.print(doc, pras);
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (fis != null) {
                    try {
                        fis.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

            return this.SUCCESS;
    }
}
