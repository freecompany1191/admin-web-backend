package com.o2osys.mng.packet.mng;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ReqSkySD0101.java
   @ReqNum      : SD01_01_V01
   @Description : 만나주문 상태(1: 배차, 2: 출발, 3: 완료, 4: 취소, 5: 배차취소) 전송
   @author      : KMS
   @since       : 2017. 10. 20.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2017. 10. 20.   KMS            최초생성

 */
@Data
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(Include.ALWAYS)
public class ReqMng {

    public ReqMng() {
    }

    /** 언어구분 (MN_SY_CODE.SY_CODE = 'LT', 0001: 한국어...) */
    @JsonProperty("language")
    private String language;

    /** 요청한 단말구분 (WCM) */
    @JsonProperty("device")
    private String device;

    /** 실제사용 정보 출력 Y/N (제휴본사, 총판, 가맹점 정보 등, API 서버를 서비스/테스트로 별도 운영하여 구분하여 세팅) */
    @JsonProperty("service_yn")
    private String serviceYn;

    /** 요청전문 */
    @JsonProperty("req_num")
    private String reqNum;

    /** 전달VALUE (DELIMETER : │) */
    @JsonProperty("value")
    private String value;

    /** 전달SESSION (DELIMETER : │) */
    @JsonProperty("session")
    private String session;

}
