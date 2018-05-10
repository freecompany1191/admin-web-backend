package com.o2osys.mng.common.excel.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.o2osys.mng.common.excel.config.ExcelConfig;

public class ExcelCommonUtil {
    // 로그
    private final Logger log = LoggerFactory.getLogger(ExcelCommonUtil.class);
    private final String TAG = ExcelCommonUtil.class.getSimpleName();
    public static int DEFAULT_COLOUMN_WIDTH = 1000;

    private Workbook wb;
    private Map<String, Object> model;
    private HttpServletResponse response;
    private CellStyle hStyle;
    private CellStyle bStyle;

    public ExcelCommonUtil(Workbook workbook, Map<String, Object> model, HttpServletResponse response) {
        this.wb = workbook;
        this.model = model;
        this.response = response;
    }

    public void createExcel() {
        setFileName(response, mapToFileName());

        createCellStyle();

        Sheet sheet = wb.createSheet();

        createHead(sheet, mapToHeadList());

        createBody(sheet, mapToBodyList());

        //log.info("response.getHeaderNames() : " + response.getHeaderNames());
        //log.info("Content-Disposition : " + response.getHeader("Content-Disposition"));
        //createFile();
    }

    private void createCellStyle() {

        /** 헤더 스타일 선언 */
        hStyle = wb.createCellStyle();
        /** 헤더 스타일 */

        hStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        hStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        hStyle.setBorderBottom(BorderStyle.THIN);
        hStyle.setBorderLeft(BorderStyle.THIN);
        hStyle.setBorderRight(BorderStyle.THIN);
        hStyle.setBorderTop(BorderStyle.THIN);
        hStyle.setAlignment(HorizontalAlignment.CENTER);

        /** 헤더 FONT 셋팅  */
        Font headerFont = wb.createFont();
        headerFont.setFontName("맑은 고딕"); // 폰트 이름
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(false);
        hStyle.setFont(headerFont);

        /** 바디 스타일 선언 */
        bStyle = wb.createCellStyle();
        /** 바디 스타일 */

        bStyle.setBorderBottom(BorderStyle.THIN);
        bStyle.setBorderLeft(BorderStyle.THIN);
        bStyle.setBorderRight(BorderStyle.THIN);
        bStyle.setBorderTop(BorderStyle.THIN);
        bStyle.setAlignment(HorizontalAlignment.LEFT);
        bStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        /** 헤더 FONT 셋팅  */
        Font bodyFont = wb.createFont();
        bodyFont.setFontName("맑은 고딕"); // 폰트 이름
        bodyFont.setFontHeightInPoints((short) 10);
        bodyFont.setBold(false);
        bStyle.setFont(bodyFont);
    }

    private void createFile() {

        //원본 파일명
        String fileName = "export.xlsx";
        log.debug(" FileName : "+fileName);

        //파일 확장자명(소문자변환)
        String fileExtension = FilenameUtils.getExtension(fileName).toLowerCase();
        log.debug(" fileExtension : "+fileExtension);

        File uploadFile;
        String uploadFileName;

        do {
            //업로드패스 (ROOT패스 + UPLOAD패스 + UPLOAD파일명)
            String uploadPath = "target/files/" + fileName;
            log.debug("uploadFilePath : "+uploadPath);

            //업로드 파일 생성
            uploadFile = new File(uploadPath);
        } while (uploadFile.exists());
        //업로드 폴더 생성
        uploadFile.getParentFile().mkdirs();

        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(uploadFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            this.wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String mapToFileName() {
        return (String) model.get(ExcelConfig.FILE_NAME);
    }

    private List<String> mapToHeadList() {
        return (List<String>) model.get(ExcelConfig.HEADER);
    }

    private List<List<String>> mapToBodyList() {
        return (List<List<String>>) model.get(ExcelConfig.BODY);
    }

    private void setFileName(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition",
                "attachment; filename=\"" + setFileExtension(fileName) + "\"");
    }

    private String setFileExtension(String fileName) {
        if ( wb instanceof XSSFWorkbook) {
            fileName += ".xlsx";
        }
        if ( wb instanceof SXSSFWorkbook) {
            fileName += ".xlsx";
        }
        if ( wb instanceof HSSFWorkbook) {
            fileName += ".xls";
        }

        //log.info("#fileName = "+fileName);

        return fileName;
    }

    private void createHead(Sheet sheet, List<String> headList) {
        createRow(sheet, headList, 0, hStyle);
    }

    private void createBody(Sheet sheet, List<List<String>> bodyList) {
        int rowSize = bodyList.size();
        for (int i = 0; i < rowSize; i++) {
            createRow(sheet, bodyList.get(i), i + 1, bStyle);
        }
    }

    private void createRow(Sheet sheet, List<String> cellList, int rowNum, CellStyle style) {
        int size = cellList.size();
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < size; i++) {
            row.createCell(i).setCellValue(cellList.get(i));
            row.getCell(i).setCellStyle(style);
        }

        for(int i = 0; i < cellList.size() ; i++) { //컬럼 width  사이즈 보기좋게 확장
            sheet.autoSizeColumn((short)i);
            //log.info("## sheet.getColumnWidth("+i+") : "+sheet.getColumnWidth(i));
            if(sheet.getColumnWidth(i) > 6000) sheet.setColumnWidth(i, 10000);
            else{
                sheet.setColumnWidth(i, (sheet.getColumnWidth(i))+DEFAULT_COLOUMN_WIDTH );  // 윗줄만으로는 컬럼의 width
            }
            //log.info("## 2sheet.getColumnWidth("+i+") : "+sheet.getColumnWidth(i));
        }
    }
}
