package com.o2osys.mng.service;

import java.io.File;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.o2osys.mng.common.service.CommonService;

@Service("FileService")
public class FileService {
    // 로그
    private final Logger log = LoggerFactory.getLogger(FileService.class);
    private final String TAG = FileService.class.getSimpleName();

    @Autowired
    CommonService commonService;

    //ROOT패스
    //D:/mnt/storege/upload/
    @Value("${file.upload.root.path}")
    String ROOT_PATH;


    /**
     * 파일업로드
     * @Method Name : fileUpload
     * @param reqFile - 업로드 파일
     * @param UPLOAD_PATH - 업로드 경로
     * @param addStr - 추가 문자열
     * @throws Exception
     */
    public String fileUpload(MultipartFile reqFile, String uploadPath, String delFileName, String ADD_STR) throws Exception {

        //파일삭제 처리
        String deleteFilePath = ROOT_PATH + delFileName;
        log.debug("deleteFilePath : "+deleteFilePath);
        boolean delChk = deleteFileExecute(deleteFilePath);
        log.debug("# delChk : "+delChk);

        //원본 파일명
        String fileName = reqFile.getOriginalFilename();
        log.debug(" FileName : "+fileName);

        //파일 확장자명(소문자변환)
        String fileExtension = FilenameUtils.getExtension(fileName).toLowerCase();
        log.debug(" fileExtension : "+fileExtension);

        File uploadFile;
        String uploadFileName;

        do {
            String dateTime = commonService.localToDate(null, "yyyyMMddhhmmss");
            log.debug("dateTime : "+dateTime);
            //업로드 파일명 생성 (현재 년월일시초+추가문자열+랜덤문자3자리)
            uploadFileName = commonService.localToDate(null, "yyyyMMddhhmmss") +ADD_STR+ RandomStringUtils.randomAlphanumeric(3) + "." + fileExtension;
            log.debug("uploadFileName : "+uploadFileName);
            //업로드패스 (ROOT패스 + UPLOAD패스 + UPLOAD파일명)
            uploadPath = ROOT_PATH + uploadPath + uploadFileName;
            log.debug("uploadFilePath : "+uploadPath);

            //업로드 파일 생성
            uploadFile = new File(uploadPath);
        } while (uploadFile.exists());
        //업로드 폴더 생성
        uploadFile.getParentFile().mkdirs();
        //파일 업로드
        reqFile.transferTo(uploadFile);

        return uploadFileName;

        //response.setAttachmentUrl("http://localhost:8080/attachments/" + destinationFileName);

    }

    /**
     * 파일 삭제
     * delete
     * @return boolean check
     */
    private boolean deleteFileExecute(String filePath) {
        boolean check = false;
        File file = new File(filePath);
        if(file.exists()) {
            check = file.delete();
        }
        return check;
    }

}
