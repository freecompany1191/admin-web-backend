package com.o2osys.mng.common.excel.domain;

import java.io.Serializable;

import org.apache.poi.ss.usermodel.Row;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class Product implements Serializable {

    private String uniqueId;

    private String name;

    private String comment;

    public static Product rowOf(Row row) {
        return Product.builder()
                .uniqueId(row.getCell(0).getStringCellValue())
                .name(row.getCell(1).getStringCellValue())
                .comment(row.getCell(2).getStringCellValue())
                .build();
    }
}
