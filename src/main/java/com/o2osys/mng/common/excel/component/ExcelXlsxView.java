package com.o2osys.mng.common.excel.component;

import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import com.o2osys.mng.common.excel.util.ExcelCommonUtil;

@Component
public class ExcelXlsxView extends AbstractXlsxView {

    @Override
    protected void buildExcelDocument(Map<String, Object> model,
            Workbook workbook,
            HttpServletRequest request,
            HttpServletResponse response) throws Exception {
        new ExcelCommonUtil(workbook, model, response).createExcel();
    }
}
