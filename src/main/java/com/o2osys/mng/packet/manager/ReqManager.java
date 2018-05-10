package com.o2osys.mng.packet.manager;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ReqManager.java
   @Description : 관리자웹사이트 전문 연결 객체
   @author      : KMS
   @since       : 2018. 1. 16.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 16.    KMS            최초생성

 */
@Data
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(Include.NON_EMPTY)
public class ReqManager {

    /** 요청 구분 login 인지 api 인지 구분 */
    @JsonProperty("REQ_TYPE")
    private String reqType;

    /** 요청_번호 */
    @JsonProperty("REQ_NUM")
    private String reqNum;

    /** 언어구분 (MN_SY_CODE.SY_CODE = 'LT', 0001: 한국어...) */
    @JsonProperty("LANGUAGE")
    private String language;

    /** 요청한 단말구분 (WCM) */
    @JsonProperty("DEVICE")
    private String device;

    /** 실제사용 정보 출력 Y/N (제휴본사, 총판, 가맹점 정보 등, API 서버를 서비스/테스트로 별도 운영하여 구분하여 세팅) */
    @JsonProperty("SERVICE_YN")
    private String serviceYn;

    /** 전달VALUE (DELIMETER : │) */
    @JsonProperty("VALUE")
    private String value;

    /* 사용자정보 (DELIMETER : │, 전달값정의: USER_ID│USER_NAME│USER_TYPE_CD│HD_CODE│BR_CODE│ST_CODE│MULTI_TYPE│) */
    @JsonProperty("INFO")
    private String info;

    /** 엑셀 다운로드 헤더 */
    @JsonProperty("HEADER")
    private String[] header;

}
