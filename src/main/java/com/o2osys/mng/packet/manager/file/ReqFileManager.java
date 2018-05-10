package com.o2osys.mng.packet.manager.file;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

/**
   @FileName  : ReqFileManager.java
   @Description : 관리자 웹 파일처리 전용 객체
   @author      : KMS
   @since       : 2018. 2. 7.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 7.    KMS            최초생성

 */
@Data
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonInclude(Include.NON_EMPTY)
public class ReqFileManager {

    /** 업로드 경로 */
    @JsonProperty("UPLOAD_PATH")
    private String uploadPath;

    /** 가맹점 코드 */
    @JsonProperty("ST_CODE")
    private String stCode;

    /** 삭제할 파일명 */
    @JsonProperty("DEL_FILE_NAME")
    private String delFileName;

}
