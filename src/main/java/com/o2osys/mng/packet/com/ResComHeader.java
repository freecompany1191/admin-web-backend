package com.o2osys.mng.packet.com;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ResAccHeader.java
   @Description : 계정 응답 공통 헤더
   @author      : KMS
   @since       : 2018. 1. 11.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 11.    KMS            최초생성

 */
@Data
@JsonInclude(JsonInclude.Include.NON_EMPTY)
public class ResComHeader {

    public ResComHeader() {
    }

    public ResComHeader(String resCode, String resMsg) {
        this.resCode = resCode;
        this.resMsg = resMsg;
    }

    /** 처리결과코드 */
    @JsonProperty("res_code")
    private String resCode;

    /** 처리결과메세지 */
    @JsonProperty("res_msg")
    private String resMsg;
}
