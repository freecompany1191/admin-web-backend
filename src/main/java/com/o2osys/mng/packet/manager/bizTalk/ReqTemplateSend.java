package com.o2osys.mng.packet.manager.bizTalk;

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
public class ReqTemplateSend {

    /** 전송 구분 등록(CREATE), 조회(SEARCH) 구분 */
    @JsonProperty("TRAN_TYPE")
    private String tranType;

    /* 사용자ID */
    @JsonProperty("USER_ID")
    private String userId;

    /** 템플릿_코드 */
    @JsonProperty("TPLT_CODE")
    private String tpltCode;

    /** 템플릿_구분 (1: 주문하기, 2: 평가하기, 3: 주문접수, 4: 출발안내) */
    @JsonProperty("TPLT_TYPE")
    private String tpltType;

    /** 템플릿_제목 */
    @JsonProperty("TPLT_TITLE")
    private String tpltTitle;

    /** 템플릿_내용 */
    @JsonProperty("TPLT_CONTENT")
    private String tpltContent;

    /** 버튼_정보 (버튼명) */
    @JsonProperty("BTN_INFO")
    private String btnInfo;

    /* 승인_상태 (0: 승인요청, 1: 승인완료, 2: 반려 3:승인요청실패) */
    @JsonProperty("APPR_STATUS")
    private String apprStatus;

}
