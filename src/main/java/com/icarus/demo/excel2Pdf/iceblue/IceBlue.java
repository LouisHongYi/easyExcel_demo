package com.icarus.demo.excel2Pdf.iceblue;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.aspose.cells.PdfSaveOptions;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.apache.commons.io.FilenameUtils;
import org.springframework.util.StopWatch;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class IceBlue {


    public static void main(String[] args) throws IOException {
        //convert2Pdf1();
        // convert2Pdf2();
        //convert2Html();
        compareFillAndConvertToPdf();
    }

    public static void convert2Pdf1(){
        //创建一个Workbook实例并加载Excel文件
        Workbook workbook = new Workbook();
        workbook.loadFromFile("C://hong//output//test2.xlsx");

        //设置转换后的PDF页面高宽适应工作表的内容大小
        workbook.getConverterSetting().setSheetFitToPage(true);

        //将生成的文档保存到指定路径
        workbook.saveToFile("C://hong//output//WorksheetToPdf1.pdf", FileFormat.PDF);

    }

    public static void compareFillAndConvertToPdf() throws IOException {
        StopWatch sw = new StopWatch("myWatch");
        sw.start("Spire start to convert excel to pdf");
        convert2Pdf2();
        sw.stop();
        System.out.println("Spire  : LastTask time (millis) = " + sw.getLastTaskTimeMillis());

        sw.start("Aspose start to convert excel to pdf");
        fillTemplateAndConvert2Pdf();
        sw.stop();
        System.out.println("Aspose: LastTask time (millis) = " + sw.getLastTaskTimeMillis());
        System.out.println(sw.prettyPrint());
        System.out.println(sw.getTotalTimeMillis());
    }

    public static void convert2Pdf2() throws IOException {

        Map<String, Object> map = new HashMap<>();
        map.put("serviceType", "PSI");
        map.put("applicant", "applicant");
        map.put("number", 65);
        map.put("unit", "N.M.R");
        map.put("country", "Durban");
        map.put("province", "4001");
        map.put("city", "South Africa");
        map.put("beneficiary", "Compagine Malagasy de textille(F001335)");
        map.put("desc", "R-Cloud-22223952");
        map.put("productDesc", "ELASTICATED SHORT WITH HRB TAPE AS TIES");
        map.put("refNo", "225385");
        map.put("quantity", 5555);
        LocalDate localDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        map.put("inspectionDate", localDate.format(formatter));
        map.put("result", "PASSED");
        map.put("text", "We, QIMA Limited, hereby certify that Compagnie Malagasy de textille(F0001335) can proceed\n" +
                "further with this Certificate and use it for delivery regarding above Products and that MRP\n" +
                "Group Apparel has decided that above goods are in good order and condition with contractual\n" +
                "specifications and as per Inspection Standards :\n" +
                "• ANSI/ASQ Standard Z.1.4-2008\n" +
                "• AQL Level I\n" +
                "• Critical Not Allowed, Major 2.5, Minor 4.0 are accepted");

        //临时excel路径
        String templatePath = "C://hong//output//Certificate_logo_template.xlsx";
        String targetExcelPath = "C://hong//output//Certificate_logo_result_spire.xlsx";
        String targetPDFPath = "C://hong//output//Certificate_logo_spire.pdf";

        // 这里使用easyExcel 填充模板生成文件，
        OutputStream os = Files.newOutputStream(Paths.get(targetExcelPath));
        File file = new File(templatePath);
        InputStream is = new FileInputStream(file);
        // InputStream is = Files.newInputStream(Paths.get(templatePath));
        ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.valueOf(FilenameUtils.getExtension(file.getName()).toUpperCase(Locale.ROOT));
        ExcelWriter excelWriter = EasyExcel.write(os)
                .withTemplate(is)// 利用模板的输出流
                .excelType(excelTypeEnum)
                .build();
        WriteSheet writeSheet = EasyExcel.writerSheet(0).build();
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();

        //创建一个Workbook实例并加载Excel文件
        Workbook workbook = new Workbook();
        workbook.loadFromFile(targetExcelPath);

        //设置转换后PDF的页面宽度适应工作表的内容宽度
       // workbook.getConverterSetting().setSheetFitToWidth(true);
        workbook.getConverterSetting().setXDpi(100);
        workbook.getConverterSetting().setSheetFitToPage(true);
        //获取第一个工作表
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //转换为PDF并将生成的文档保存到指定路径
        worksheet.saveToPdf("C://hong//output//Certificate_logo_result_spire.pdf");
    }

    public static void convert2Html() {
        //创建一个Workbook实例并加载Excel文件
        Workbook workbook = new Workbook();
        workbook.loadFromFile("C://hong//output//test2.xlsx");

        //设置转换后PDF的页面宽度适应工作表的内容宽度
        workbook.getConverterSetting().setSheetFitToWidth(true);

        //获取第一个工作表
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //转换为PDF并将生成的文档保存到指定路径
        worksheet.saveToHtml("C://hong//output//WorksheetToHtml.html");
    }


    public static void fillTemplateAndConvert2Pdf() {
        try {
            Map<String, Object> map = new HashMap<>();
            map.put("serviceType", "PSI");
            map.put("applicant", "applicant");
            map.put("number", 65);
            map.put("unit", "N.M.R");
            map.put("country", "Durban");
            map.put("province", "4001");
            map.put("city", "South Africa");
            map.put("beneficiary", "Compagine Malagasy de textille(F001335)");
            map.put("desc", "R-Cloud-22223952");
            map.put("productDesc", "ELASTICATED SHORT WITH HRB TAPE AS TIES");
            map.put("refNo", "225385");
            map.put("quantity", 5555);
            LocalDate localDate = LocalDate.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
            map.put("inspectionDate", localDate.format(formatter));
            map.put("result", "PASSED");
            map.put("text", "We, QIMA Limited, hereby certify that Compagnie Malagasy de textille(F0001335) can proceed\n" +
                    "further with this Certificate and use it for delivery regarding above Products and that MRP\n" +
                    "Group Apparel has decided that above goods are in good order and condition with contractual\n" +
                    "specifications and as per Inspection Standards :\n" +
                    "• ANSI/ASQ Standard Z.1.4-2008\n" +
                    "• AQL Level I\n" +
                    "• Critical Not Allowed, Major 2.5, Minor 4.0 are accepted");

            //临时excel路径
            String templatePath = "C://hong//output//Certificate_logo_template.xlsx";
            String targetExcelPath = "C://hong//output//Certificate_logo_result_aspose.xlsx";
            String targetPDFPath = "C://hong//output//Certificate_logo_result_aspose.pdf";

            // 这里使用easyExcel 填充模板生成文件，
            OutputStream os = Files.newOutputStream(Paths.get(targetExcelPath));
            File file = new File(templatePath);
            InputStream is = new FileInputStream(file);
            // InputStream is = Files.newInputStream(Paths.get(templatePath));
            ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.valueOf(FilenameUtils.getExtension(file.getName()).toUpperCase(Locale.ROOT));
            ExcelWriter excelWriter = EasyExcel.write(os)
                    .withTemplate(is)// 利用模板的输出流
                    .excelType(excelTypeEnum)
                    .build();
            WriteSheet writeSheet = EasyExcel.writerSheet(0).build();
            excelWriter.fill(map, writeSheet);
            excelWriter.finish();

            /*Workbook wb = new Workbook(targetExcelPath);
            wb.save(targetPDFPath, SaveFormat.PDF);*/

            FileOutputStream fileOS = new FileOutputStream(targetPDFPath);
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(targetExcelPath);
            //把内容放在一张PDF 页面上；
            pdfSaveOptions.setOnePagePerSheet(true);

            workbook.save(fileOS, pdfSaveOptions);
            fileOS.flush();
            fileOS.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
