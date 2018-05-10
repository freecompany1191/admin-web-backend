package com.o2osys.mng.common.service;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Component;

/**
   @FileName  : CommonService.java
   @Description : 공통 서비스
   @author      : KMS
   @since       : 2017. 7. 21.
   @version     : 1.0

   @개정이력

   수정일           수정자         수정내용
   -----------      ---------      -------------------------------
   2017. 7. 21.     KMS            최초생성

 */
@Component
public class CommonService {
    // 로그
    private final Logger LOGGER = LoggerFactory.getLogger(CommonService.class);
    private final String TAG = CommonService.class.getSimpleName();

    @Autowired
    MessageSource messageSource;

    /**
     * 에러로그를 뿌려준다
     * @param e
     */
    public void errorLog(String errorTag, Exception e) {
        LOGGER.error(errorTag, e);
        LOGGER.error(e.getMessage());
        LOGGER.error(e.toString());
        StackTraceElement[] ste = e.getStackTrace();
        for(int i = 0; i < ste.length; i++)
        {
            LOGGER.error(String.valueOf(ste[i]));
        }
    }

    public String getMessage(String code){

        return messageSource.getMessage(code,null,Locale.KOREA);

    }

    /**
     * 날짜 형식 변경
     * yyyyMMddHHmmss -> yyyy-MM-dd HH:mm:ss
     * @Method Name : strToDate
     * @param strDate
     * @return
     */
    public static String localToDate(LocalDateTime toDate, String pattern) throws Exception{

        if(toDate == null){
            toDate = LocalDateTime.now();
        }

        DateTimeFormatter formater;

        if(pattern != null) formater = DateTimeFormatter.ofPattern(pattern, Locale.KOREA);
        else                formater = DateTimeFormatter.ofPattern("yyyy-MM-dd a hh:mm:ss", Locale.KOREA);

        String formatStr = formater.format(toDate);

        return formatStr;
    }

}
