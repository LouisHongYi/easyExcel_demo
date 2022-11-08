package com.icarus.demo.excel2Pdf.aspose;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.aspose.cells.*;
import com.icarus.demo.util.EncodeUtils;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.commons.io.FilenameUtils;

import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class Aspose {

    public static void main(String[] args) {
        try {

            //convert2Pdf();
            // excel2pdf("C://hong//output//test2.xlsx", "C://hong//output//output-xlsx.pdf");

            fillTemplateAndConvert2Pdf();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void convert2Pdf() throws Exception {
        String dir = "C://hong//output//";
        File file = new File(dir + "test2.xlsx");

        Workbook wb = new Workbook(file.getPath());
        wb.save(dir + "output.pdf", SaveFormat.PDF);
        wb.save(dir + "output.xps", SaveFormat.XPS);
        wb.save(dir + "output.html", SaveFormat.HTML);
        //targetFile.delete();
    }

    /**
     * 获取license 去除水印
     *
     * @return
     */
    public static boolean getLicense() {
        try {
            String license = "PExpY";
            InputStream is = BaseToInputStream(license);
            License aposeLic = new License();
            aposeLic.setLicense(is);
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * base64转inputStream
     *
     * @param base64string
     * @return
     */
    private static InputStream BaseToInputStream(String base64string) {
        ByteArrayInputStream stream = null;
        try {

            byte[] bytes1 = EncodeUtils.base64Decode(base64string);
            stream = new ByteArrayInputStream(bytes1);
        } catch (Exception e) {
            // TODO: handle exception
        }
        return stream;
    }

    /**
     * excel 转为pdf 输出。
     *
     * @param sourceFilePath excel文件
     * @param desFilePath    pad 输出文件目录
     */
    public static void excel2pdf(String sourceFilePath, String desFilePath) {
        // 验证License 若不验证则转化出的pdf文档会有水印产生
/*        if (!getLicense()) {
            return;
        }*/
        try {
            /*LoadOptions loadOptions = new LoadOptions();
            loadOptions.setRegion(CountryCode.CHINA);

            Workbook wb = new Workbook(sourceFilePath, loadOptions);*/

            Workbook wb = new Workbook(sourceFilePath);
            FileOutputStream fileOS = new FileOutputStream(desFilePath);
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            AutoFitterOptions options = new AutoFitterOptions();
            options.setAutoFitMergedCellsType(AutoFitMergedCellsType.EACH_LINE);
            wb.getWorksheets().get(0).autoFitRows(options);


            //把内容放在一张PDF 页面上；
            pdfSaveOptions.setOnePagePerSheet(true);
            int[] autoDrawSheets = {3};
            //当excel中对应的sheet页宽度太大时，在PDF中会拆断并分页。此处等比缩放。
            autoDraw(wb, autoDrawSheets);
            int[] showSheets = {0};
            //隐藏workbook中不需要的sheet页。
            printSheetPage(wb, showSheets);

            wb.save(fileOS, pdfSaveOptions);
            fileOS.flush();
            fileOS.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 设置打印的sheet 自动拉伸比例
     *
     * @param wb
     * @param page 自动拉伸的页的sheet数组
     */
    public static void autoDraw(Workbook wb, int[] page) {
        if (null != page && page.length > 0) {
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).getHorizontalPageBreaks().clear();
                wb.getWorksheets().get(i).getVerticalPageBreaks().clear();
            }
        }
    }


    /**
     * 隐藏workbook中不需要的sheet页。
     *
     * @param wb
     * @param page 显示页的sheet数组
     */
    public static void printSheetPage(Workbook wb, int[] page) {
        for (int i = 1; i < wb.getWorksheets().getCount(); i++) {
            wb.getWorksheets().get(i).setVisible(false);
        }
        if (null == page || page.length == 0) {
            wb.getWorksheets().get(0).setVisible(true);
        } else {
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).setVisible(true);
            }
        }
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
            String targetExcelPath = "C://hong//output//Certificate_logo_result.xlsx";
            String targetPDFPath = "C://hong//output//Certificate_logo.pdf";

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

            Workbook workbook = new Workbook(targetExcelPath);
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
