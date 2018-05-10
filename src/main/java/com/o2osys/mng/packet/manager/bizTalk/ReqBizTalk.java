package com.o2osys.mng.packet.manager.bizTalk;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ReqBizTalk.java
   @Description : 카카오 비즈톡 전문 전송 객체
   @author      : KMS
   @since       : 2018. 2. 1.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 1.    KMS            최초생성

 */
@Data
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(Include.NON_EMPTY)
public class ReqBizTalk {

    /** 발신 프로필 키(senderKeyType이 G인경우발신 프로필그룹키) */
    @JsonProperty("senderKey")
    private String senderKey;

    /** 템플릿 코드 */
    @JsonProperty("templateCode")
    private String templateCode;

    /** 템플릿 이름 */
    @JsonProperty("templateName")
    private String templateName;

    /** 템플릿 내용 */
    @JsonProperty("templateContent")
    private String templateContent;

    /** 발신 프로필 키 타입 (G:그룹 , S:발신프로필(default)) */
    @JsonProperty("senderKeyType")
    private String senderKeyType;

    /** 버튼타입 (N:없음(default), DS:배송조회, C:자유설정) */
    @JsonProperty("buttonType")
    private String buttonType;

    /** 버튼이름 (버튼타입 C인 경우 필수로 입력해야함) */
    @JsonProperty("buttonName")
    private String buttonName;

    /** 버튼 링크 주소(버튼타입 C인 경우 필수로 입력해야함) */
    @JsonProperty("buttonUrl")
    private String buttonUrl;

}
