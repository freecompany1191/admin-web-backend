package com.o2osys.mng.common.util;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Random;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.util.StringUtils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.o2osys.mng.common.constants.Const;

/**
   @FileName  : CommonUtil.java
   @Description : 공통유틸
   @author      : KMS
   @since       : 2017. 7. 21.
   @version     : 1.0

   @개정이력

   수정일           수정자         수정내용
   -----------      ---------      -------------------------------
   2017. 7. 21.     KMS            최초생성

 */
public class CommonUtils {

    /**
     * Object -> JSON 형식으로 변환
     *
     * @Method Name : jsonStringFromObject
     * @param object
     * @return
     * @throws JsonProcessingException
     */
    public static String jsonStringFromObject(Object object) throws JsonProcessingException {

        if(object == null){
            return null;
        }

        ObjectMapper mapper = new ObjectMapper();
        return mapper.writeValueAsString(object);
    }

    /**
     * 날짜 형식 변경
     * yyyyMMddHHmmss -> yyyy-MM-dd HH:mm:ss
     * @Method Name : strToDate
     * @param strDate
     * @return
     */
    public static String strToDate(String strDate) {

        if(StringUtils.isEmpty(strDate)){
            return null;
        }

        DateTimeFormatter pattern = DateTimeFormatter.ofPattern("yyyyMMddHHmmss", Locale.KOREA);
        LocalDateTime toDate = LocalDateTime.parse(strDate,pattern);

        pattern = DateTimeFormatter.ofPattern("yyyy.MM.dd HH;mm;ss", Locale.KOREA);
        String formatStr = pattern.format(toDate);

        return formatStr;
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
            return null;
        }

        DateTimeFormatter formater;

        if(pattern != null) formater = DateTimeFormatter.ofPattern(pattern, Locale.KOREA);
        else                formater = DateTimeFormatter.ofPattern("yyyy-MM-dd hh:mm:ss", Locale.KOREA);

        String formatStr = formater.format(toDate);

        return formatStr;
    }

    /**
     * 배달대행 공통 전문 헤더 설정
     *
     * @Method Name : getHeader
     * @return
     */
    public static HttpHeaders getHeader(){

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON_UTF8);

        return headers;
    }

    /**
     * 현재시간 (YYYYMMDDHH24MISS)
     *
     * @Method Name : getNowTime
     * @return
     */
    public static String getNowTime(){

        DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

        LocalDateTime now = LocalDateTime.now();
        String nowTime = now.format(format);

        return nowTime;
    }

    /**
     * 배달대행 전문추적번호 생성(YYYYMMDDHH24MISS + Random 자릿수 지정)
     *
     * @Method Name : getTraceNo
     * @return
     */
    public static String getTraceNo(int num){

        DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

        //입력된 숫자만큼 Random 자릿수 셋팅
        StringBuilder sb = new StringBuilder();
        for (int i=0;i<num;i++) {
            sb.append("9");
        }
        int cnt = Integer.valueOf(sb.toString());

        Random rn = new Random();
        String rnum = String.format("%0"+num+"d", rn.nextInt(cnt));

        LocalDateTime now = LocalDateTime.now();
        String nowTime = now.format(format);
        String traceNo = nowTime + rnum;

        return traceNo;
    }


    /**
     * 코드 컨버팅
     * @Method Name : Transfer
     * @return
     */
    public static String codeConv(String name, String code) {

        String convValue = "";

        if(StringUtils.isEmpty(code)){
            return convValue;
        }

        switch(name){

            /** 요청 타입 */
            case "REQ_TYPE" :

                switch(code){
                    case Const.REQ.TYPE.API :
                        convValue = "";
                        break;
                    case Const.REQ.TYPE.API1 :
                        convValue = "_1";
                        break;
                    case Const.REQ.TYPE.API2 :
                        convValue = "_2";
                        break;
                    case Const.REQ.TYPE.API3 :
                        convValue = "_3";
                        break;
                    case Const.REQ.TYPE.API4 :
                        convValue = "_4";
                        break;
                }
                break;


                /** 카카오 비즈톡 응답 메세지 */
            case "BIZ_TALK_RES" :

                switch(code){
                    case Const.BIZ_TALK.RES.CODE.OK :
                        convValue = Const.BIZ_TALK.RES.MSG.OK;
                        break;
                    case Const.BIZ_TALK.RES.CODE.NO_AUTH :
                        convValue = Const.BIZ_TALK.RES.MSG.NO_AUTH;
                        break;
                    case Const.BIZ_TALK.RES.CODE.PARAM_ERROR :
                        convValue = Const.BIZ_TALK.RES.MSG.PARAM_ERROR;
                        break;
                    case Const.BIZ_TALK.RES.CODE.DUPLE_CODE :
                        convValue = Const.BIZ_TALK.RES.MSG.DUPLE_CODE;
                        break;
                    case Const.BIZ_TALK.RES.CODE.DUPLE_NAME :
                        convValue = Const.BIZ_TALK.RES.MSG.DUPLE_NAME;
                        break;
                    case Const.BIZ_TALK.RES.CODE.BYTE_ERROR :
                        convValue = Const.BIZ_TALK.RES.MSG.BYTE_ERROR;
                        break;
                    case Const.BIZ_TALK.RES.CODE.NO_SENDER :
                        convValue = Const.BIZ_TALK.RES.MSG.NO_SENDER;
                        break;
                    case Const.BIZ_TALK.RES.CODE.NO_REQ_DATA :
                        convValue = Const.BIZ_TALK.RES.MSG.NO_REQ_DATA;
                        break;
                    case Const.BIZ_TALK.RES.CODE.NO_STATUS :
                        convValue = Const.BIZ_TALK.RES.MSG.NO_STATUS;
                        break;
                    case Const.BIZ_TALK.RES.CODE.BIZ_AUTH_KEY_ERROR :
                        convValue = Const.BIZ_TALK.RES.MSG.BIZ_AUTH_KEY_ERROR;
                        break;
                    case Const.BIZ_TALK.RES.CODE.BIZ_API_ERROR :
                        convValue = Const.BIZ_TALK.RES.MSG.BIZ_API_ERROR;
                        break;
                    case Const.BIZ_TALK.RES.CODE.NOT_CONNECT :
                        convValue = Const.BIZ_TALK.RES.MSG.NOT_CONNECT;
                        break;
                    case Const.BIZ_TALK.RES.CODE.SYSTEM_ERROR :
                        convValue = Const.BIZ_TALK.RES.MSG.SYSTEM_ERROR;
                        break;
                }
                break;

        }
        return convValue;
    }

}
