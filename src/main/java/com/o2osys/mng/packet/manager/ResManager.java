package com.o2osys.mng.packet.manager;

import java.util.ArrayList;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ResManager.java
   @Description : 관리자웹사이트 전문 연결 응답 객체
   @author      : KMS
   @since       : 2018. 1. 16.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 16.    KMS            최초생성

 */
@Data
@JsonInclude(Include.ALWAYS)
public class ResManager {

    /** 결과값1 */
    @JsonProperty("out_CODE")
    private String outCode;

    /** 결과값1 */
    @JsonProperty("out_MSG")
    private String outMsg;

    /** 결과값1 */
    @JsonProperty("out_VALUE")
    private String outValue;

    /** 결과값1 */
    @JsonProperty("out_ROW1")
    private ArrayList outRow1;

    /** 결과값2 */
    @JsonProperty("out_ROW2")
    private ArrayList outRow2;

    /** 결과값3 */
    @JsonProperty("out_ROW3")
    private ArrayList outRow3;

}