package com.o2osys.mng.entity;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : GrpTpltVO.java
   @Description : 그룹 템플릿 객체
   @author      : KMS
   @since       : 2018. 2. 1.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 1.    KMS            최초생성

 */
@Data
@JsonInclude(Include.ALWAYS)
public class GrpTpltVO {

    /** 템플릿_코드 */
    @JsonProperty("TPLT_CODE")
    private String tpltCode;

    /** 데이터_상태 (0: 삭제, 1: 정상) */
    @JsonProperty("DATA_STATUS")
    private String dataStatus;

    /** 승인_상태 (0: 승인요청, 1: 승인완료, 2: 반려 3:승인요청실패) */
    @JsonProperty("APPR_STATUS")
    private String apprStatus;

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

    /** 서비스_구분 (1: 만나, 2: BBQ) */
    @JsonProperty("SERVICE_TYPE")
    private String serviceType;

    /** 센더_키_구분 (G: 그룹, S: 개별) */
    @JsonProperty("SENDER_KEY_TYPE")
    private String senderKeyType;

    /** 센더_키 */
    @JsonProperty("SENDER_KEY")
    private String senderKey;

    /** 반려 메세지 */
    @JsonProperty("ERROR_MSG")
    private String errorMsg;

    /** 사용자ID */
    @JsonProperty("USER_ID")
    private String userId;

}
