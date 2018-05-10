package com.o2osys.mng.packet.manager.bizTalk;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.o2osys.mng.packet.com.ResCommon;

import lombok.Data;

/**
   @FileName  : ResBizTalk.java
   @Description : 카카오 비즈톡 전문 응답 객체
   @author      : KMS
   @since       : 2018. 2. 1.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 1.    KMS            최초생성

 */
@Data
@JsonInclude(Include.NON_EMPTY)
public class ResBizTalk extends ResCommon {

    /** 결과 코드 */
    @JsonProperty("code")
    private String code;

    /** 결과 코드 */
    @JsonProperty("message")
    private String message;

    /** 결과 데이터 */
    @JsonProperty("data")
    private ResBizTalkData data;

    @Data
    @JsonIgnoreProperties(ignoreUnknown = true)
    @JsonInclude(Include.NON_EMPTY)
    public static class ResBizTalkData {

        /** 발신프로필키 */
        @JsonProperty("senderKey")
        private String senderKey;

        /** 발신프로필타입(S:기본, G:발신프로필그룹) */
        @JsonProperty("senderKeyType")
        private String senderKeyType;

        /** 템플릿코드 */
        @JsonProperty("templateCode")
        private String templateCode;

        /** 템플릿이름 */
        @JsonProperty("templateName")
        private String templateName;

        /** 템플릿내용 */
        @JsonProperty("templateContent")
        private String templateContent;

        /** 검사상태(REG:접수, APL:등록, INS:검수중, COM:승인, REJ: 반려) */
        @JsonProperty("inspectionStatus")
        private String inspectionStatus;

        /** 등록일 */
        @JsonProperty("createdAt")
        private String createdAt;

        /** 수정일 */
        @JsonProperty("modifiedAt")
        private String modifiedAt;

        /** 템플릿 삭제(S:중단, A:정상, R:대기(발송전)) */
        @JsonProperty("status")
        private String status;

        /** 반려사유 */
        @JsonProperty("comments")
        private String comments;

        /** 버튼타입 (N:없음(default), DS:배송조회, C:자유설정) */
        @JsonProperty("buttonType")
        private String buttonType;

        /** 버튼이름 */
        @JsonProperty("buttonName")
        private String buttonName;

        /** 버튼 링크 주소 */
        @JsonProperty("buttonUrl")
        private String buttonUrl;

    }


}