package com.o2osys.mng.entity;

import java.util.Date;

import com.fasterxml.jackson.annotation.JsonFormat;

import lombok.Data;

/**
   @FileName  : MnGrSt.java
   @Description :
   @author      : KMS
   @since       : 2018. 1. 8.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 8.    KMS            최초생성

 */
@Data
public class MnGrSt {
    /** 가맹점_코드 **/
    String stCode;
    /** 입력_일시 */
    @JsonFormat(pattern="yyyy-MM-dd HH:mm:ss")
    Date putDate;
    /** 입력_사용자_ID **/
    String putUserId;
    /** 가맹점_이름 **/
    String StName;
    /** 가맹점_이름_앱표시명 (고객앱 가맹점명) **/
    String stNameApp;
    /** 가맹점_대표자명 **/
    String stOwner;
    /** 가맹점_전화번호 **/
    String stTel;
}
