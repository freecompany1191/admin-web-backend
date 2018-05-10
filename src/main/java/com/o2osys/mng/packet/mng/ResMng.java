package com.o2osys.mng.packet.mng;

import java.util.ArrayList;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.o2osys.mng.packet.com.ResCommon;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ResInputOrder.java
   @Description : BAROGO API 3. 배송요청(API_DVRY_INPUT_ORDER) 응답 객체
   @author      : KMS
   @since       : 2017. 9. 7.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2017. 9. 7.     KMS            최초생성

 */
@Data
@JsonInclude(Include.NON_EMPTY)
public class ResMng {

    public ResMng() {
    }

    public ResMng(ResMngHeader header, ResMngBody body) {
        this.header = header;
        this.body = body;
    }

    @JsonProperty("header")
    private ResMngHeader header;

    @JsonProperty("body")
    private ResMngBody body;

    @Data
    @JsonInclude(Include.NON_EMPTY)
    public static class ResMngBody extends ResCommon {

        public ResMngBody() {
        }

        /** 결과값1 */
        @JsonProperty("out_row1")
        private ArrayList outRow1;

        /** 결과값2 */
        @JsonProperty("out_row2")
        private ArrayList outRow2;

        /** 결과값3 */
        @JsonProperty("out_row3")
        private ArrayList outRow3;

    }

}