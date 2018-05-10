package com.o2osys.mng.service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.o2osys.mng.common.constants.Const;
import com.o2osys.mng.common.constants.Const.RES;
import com.o2osys.mng.common.constants.Define.Lang;
import com.o2osys.mng.common.constants.Define.SYSTEM;
import com.o2osys.mng.common.exception.AlreadyCompleteException;
import com.o2osys.mng.common.exception.ProcessException;
import com.o2osys.mng.common.exception.RequestException;
import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.mapper.ManagerMapper;
import com.o2osys.mng.packet.manager.ReqManager;
import com.o2osys.mng.packet.manager.ResManager;
import com.o2osys.mng.packet.manager.bizTalk.ReqTemplateSend;
import com.o2osys.mng.packet.manager.file.ReqFileManager;;

/**
   @FileName  : ManagerService.java
   @Description :
   @author      : KMS
   @since       : 2018. 2. 27.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 27.    KMS            최초생성

 */
@Service("ManagerService")
public class ManagerService {
    // 로그
    private final Logger log = LoggerFactory.getLogger(ManagerService.class);
    private final String TAG = ManagerService.class.getSimpleName();

    @Autowired
    ManagerMapper managerMapper;

    @Autowired
    FileService fileService;

    @Autowired
    BizTalkService bizTalkService;

    @Autowired
    CommonService commonService;


    /**
     * 관리자웹사이트 API 전문 호출 로직
     * @Method Name : callApi
     * @param reqManager
     * @return
     * @throws Exception
     */
    public ResManager callApi(ReqManager reqManager, ReqFileManager reqFileManager, MultipartFile reqFile, String apiType) throws Exception {

        ResManager resManager = new ResManager();

        try {

            if("fileApi".equals(apiType)){
                return callApis(reqFileManager, reqFile);
            }else {
                return callApis(reqManager);
            }

        }

        catch (RequestException e) {
            /** Request 파라미터 오류(필수누락 등) - 0001 */
            log.error(e.getMessage());
            resManager.setOutCode(RES.CODE.FAIL);
            resManager.setOutMsg(e.getMessage());
            resManager.setOutRow1(new ArrayList());
            resManager.setOutRow2(new ArrayList());
            resManager.setOutRow3(new ArrayList());

            return resManager;

        } catch (ProcessException e) {
            /** 처리거절(처리 불가능 상황) - 0002 */
            log.error(e.getMessage());
            resManager.setOutCode(RES.CODE.FAIL);
            resManager.setOutMsg(e.getMessage());
            resManager.setOutRow1(new ArrayList());
            resManager.setOutRow2(new ArrayList());
            resManager.setOutRow3(new ArrayList());

            return resManager;

        } catch (AlreadyCompleteException e) {
            /** 기 처리상태(이미 처리된 상태) - 0003 */
            log.error(e.getMessage());
            resManager.setOutCode(RES.CODE.FAIL);
            resManager.setOutMsg(e.getMessage());
            resManager.setOutRow1(new ArrayList());
            resManager.setOutRow2(new ArrayList());
            resManager.setOutRow3(new ArrayList());

            return resManager;

        } catch (Exception e) {
            /** 시스템 오류(정의되지 않은 오류) - 9999 */
            commonService.errorLog(TAG, e);
            resManager.setOutCode(RES.CODE.FAIL);
            resManager.setOutMsg(RES.MSG.SYSTEM_ERROR);
            resManager.setOutRow1(new ArrayList());
            resManager.setOutRow2(new ArrayList());
            resManager.setOutRow3(new ArrayList());

            return resManager;
        }

    }

    /**
     * 관리자웹사이트 파일 처리 전용 서비스
     * @Method Name : callApis
     * @param reqManager
     * @return
     * @throws Exception
     */
    public ResManager callApis(ReqFileManager reqFileManager, MultipartFile reqFile) throws Exception {

        ResManager resManager = new ResManager();

        if(reqFileManager == null){
            throw new RequestException(commonService.getMessage("error.not.reqvalue") + " : is Not Empty param");
        }

        if(reqFile == null){
            throw new RequestException(commonService.getMessage("error.not.reqvalue") + " : is Not Exists File");
        }

        if(StringUtils.isEmpty(reqFileManager.getUploadPath())){
            throw new RequestException(commonService.getMessage("error.not.reqvalue") + " : is Not Empty UPLOAD_PATH");
        }

        if(StringUtils.isEmpty(reqFileManager.getStCode())){
            throw new RequestException(commonService.getMessage("error.not.reqvalue") + " : is Not Empty ST_CODE");
        }

        //업로드 경로
        String UPLOAD_PATH = reqFileManager.getUploadPath();
        //삭제할 문자열
        String delFileName = reqFileManager.getDelFileName();
        //가맹점 코드
        String stCode = reqFileManager.getStCode();

        //ST_CODE 뒤 3자리 문자열
        String ST_CODE_SUB_PATH = stCode.substring(4)+"/";
        //ST_CODE
        String ST_CODE_PATH = stCode+"/";

        //업로드 경로 조합
        UPLOAD_PATH = UPLOAD_PATH + ST_CODE_SUB_PATH + ST_CODE_PATH;

        String uploadFileName = null;
        //파일이 있고 업로드 경로가 있을때만 업로드
        if(reqFile != null && UPLOAD_PATH != null){
            uploadFileName = fileService.fileUpload(reqFile, UPLOAD_PATH, delFileName, "");
        }

        Map<String, Object> map = new HashMap<String, Object>();
        map.put("UPLOAD_PATH", UPLOAD_PATH);
        map.put("FILE_NAME", uploadFileName);

        ArrayList<Map<String, Object>> outRow1 = new ArrayList<Map<String, Object>>();
        outRow1.add(map);

        resManager.setOutCode(RES.CODE.OK);
        resManager.setOutRow1(outRow1);

        return resManager;

    }

    /**
     * 관리자웹사이트 전문 연결 통합 API 호출 로직
     * @Method Name : callApis
     * @param reqManager
     * @return
     * @throws Exception
     */
    public ResManager callApis(ReqManager reqManager) throws Exception {

        ResManager resManager = new ResManager();

        if(reqManager == null){
            throw new RequestException(commonService.getMessage("error.not.reqvalue"));
        }

        if(StringUtils.isEmpty(reqManager.getValue())){
            throw new RequestException(commonService.getMessage("error.not.reqvalue"));
        }

        if(StringUtils.isEmpty(reqManager.getLanguage())) reqManager.setLanguage(Lang.KO);
        if(StringUtils.isEmpty(reqManager.getDevice())) reqManager.setDevice(SYSTEM.TYPE);
        if(StringUtils.isEmpty(reqManager.getServiceYn())) reqManager.setServiceYn("N");


        if(StringUtils.isEmpty(reqManager.getReqNum())){
            throw new RequestException(RES.MSG.REQ_ERROR);
        }

        //전문타입에 따른 구분
        switch(reqManager.getReqType()) {
            case Const.REQ.TYPE.LOGIN : //관리자웹사이트 로그인 API 호출 로직
                resManager = spSyApiWeb(reqManager);
                break;
            case Const.REQ.TYPE.API : //관리웹사이트 전문 연결 API 처리 로직
            case Const.REQ.TYPE.API1 : //관리웹사이트 전문 연결 API1 처리 로직
            case Const.REQ.TYPE.API2 : //관리웹사이트 전문 연결 API2 처리 로직
            case Const.REQ.TYPE.API3 : //관리웹사이트 전문 연결 API3 처리 로직
            case Const.REQ.TYPE.API4 : //관리웹사이트 전문 연결 API4 처리 로직
                resManager = (ResManager) spSyApiAdmWeb(reqManager, false);

                //요청 전문이 템플릿 등록이면 비즈톡 템플릿 등록 요청 서비스를 호출한다

                if(Const.REQ.TYPE.API3.equals(reqManager.getReqType()) && "AW03_02_V01".equals(reqManager.getReqNum())){

                    ReqTemplateSend reqTemplateSend = new ReqTemplateSend();

                    if(reqManager.getValue() != null){

                        String tpltCode = resManager.getOutValue();

                        String[] strArr = reqManager.getValue().split("│");

                        if(strArr.length >= 4){

                            reqTemplateSend.setTranType("CREATE");
                            reqTemplateSend.setTpltCode(tpltCode);
                            reqTemplateSend.setTpltType(strArr[0]);
                            reqTemplateSend.setTpltTitle(strArr[1]);
                            reqTemplateSend.setTpltContent(strArr[2]);
                            reqTemplateSend.setBtnInfo(strArr[3]);

                            //비즈톡 템플릿 등록 요청 서비스 호출
                            bizTalkService.sendTemplate(reqTemplateSend);

                        }else {
                            throw new RequestException(commonService.getMessage("error.not.reqvalue"));
                        }
                    }else {
                        throw new RequestException(commonService.getMessage("error.not.reqvalue"));
                    }

                }

                break;
            default :
                throw new RequestException(commonService.getMessage("error.not.reqtype"));
        }

        return resManager;

    }

    /**
     * 관리자웹사이트 전문 연결 Excel API 호출 로직
     * @Method Name : callExcelApis
     * @param reqManager
     * @return Map
     * @throws Exception
     */
    public Map<String, Object> callExcelApis(ReqManager reqManager) throws Exception {

        Map<String, Object> resultMap = new HashMap<String, Object>();

        if(reqManager == null){
            throw new RequestException(commonService.getMessage("error.not.reqvalue"));
        }

        if(StringUtils.isEmpty(reqManager.getValue())){
            throw new RequestException(commonService.getMessage("error.not.reqvalue"));
        }

        if(StringUtils.isEmpty(reqManager.getLanguage())) reqManager.setLanguage(Lang.KO);
        if(StringUtils.isEmpty(reqManager.getDevice())) reqManager.setDevice(SYSTEM.TYPE);
        if(StringUtils.isEmpty(reqManager.getServiceYn())) reqManager.setServiceYn("N");


        if(StringUtils.isEmpty(reqManager.getReqNum())){
            throw new RequestException(RES.MSG.REQ_ERROR);
        }

        //전문타입에 따른 구분
        switch(reqManager.getReqType()) {
            case Const.REQ.TYPE.API : //관리웹사이트 전문 연결 API 처리 로직
            case Const.REQ.TYPE.API1 : //관리웹사이트 전문 연결 API1 처리 로직
            case Const.REQ.TYPE.API2 : //관리웹사이트 전문 연결 API2 처리 로직
            case Const.REQ.TYPE.API3 : //관리웹사이트 전문 연결 API3 처리 로직
            case Const.REQ.TYPE.API4 : //관리웹사이트 전문 연결 API4 처리 로직
                resultMap = (Map<String, Object>) spSyApiAdmWeb(reqManager, true);
                break;
            default :
                throw new RequestException(commonService.getMessage("error.not.reqtype"));
        }

        return resultMap;

    }


    /**
     * 관리자웹사이트 로그인 API 처리 로직
     * @Method Name : spSyApiWeb
     * @param reqManager
     * @return
     * @throws Exception
     */
    public ResManager spSyApiWeb(ReqManager reqManager) throws Exception {
        Map<String, Object> map = new HashMap<String, Object>();

        /** 요청전문 */
        map.put("in_REQ_NUM", reqManager.getReqNum());
        /** 언어구분 (MN_SY_CODE.SY_CODE = 'LT', 0001:한국어...) */
        map.put("in_LANGUAGE", reqManager.getLanguage());
        /** 요청한 단말구분 */
        map.put("in_DEVICE", reqManager.getDevice());
        /** 실제사용 정보 출력 Y/N (제휴본사, 총판, 가맹점 정보 등, API 서버를 서비스/테스트로 별도 운영하여 구분하여 세팅) */
        map.put("in_SERVICE_YN", reqManager.getServiceYn());

        /** 전달VALUE (DELIMETER : │) */
        map.put("in_VALUE", reqManager.getValue());

        /** 사용자정보 (DELIMETER : │, 전달값정의: USER_ID│USER_NAME│USER_TYPE_CD│HD_CODE│BR_CODE│ST_CODE│MULTI_TYPE│) */
        map.put("in_INFO", reqManager.getInfo());

        log.debug("[CALL SP_SY_API_WEB(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+") JSON : "+CommonUtils.jsonStringFromObject(map));

        //관리자웹사이트 로그인 API 호출
        managerMapper.spSyApiWeb(map);

        log.info("[SP_SY_API_WEB(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE : "+map.get("out_CODE")+", out_MSG : "+map.get("out_MSG"));

        if(map.get("out_CODE") == null){

            throw new Exception("[SP_SY_API_WEB(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE is Null : "+map.get("out_CODE"));

        }else if(!NumberUtils.isDigits( String.valueOf(map.get("out_CODE")).trim() )){

            throw new Exception("[SP_SY_API_WEB(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE is Not Number : "+map.get("out_CODE"));

        }

        ResManager resManager = new ResManager();

        resManager.setOutCode(String.valueOf(map.get("out_CODE")));
        log.debug("[SP_SY_API_WEB] out_CODE : "+CommonUtils.jsonStringFromObject(map.get("out_CODE")));
        resManager.setOutMsg(String.valueOf(map.get("out_MSG")));
        log.debug("[SP_SY_API_WEB] out_MSG : "+CommonUtils.jsonStringFromObject(map.get("out_MSG")));
        resManager.setOutValue(String.valueOf(map.get("out_VALUE")));
        log.debug("[SP_SY_API_WEB] out_VALUE : "+CommonUtils.jsonStringFromObject(map.get("out_VALUE")));
        resManager.setOutRow1(convertMap(map.get("out_ROW1"), "SP_SY_API_WEB"));
        log.debug("[SP_SY_API_WEB] out_ROW1 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW1")));
        resManager.setOutRow2(convertMap(map.get("out_ROW2"), "SP_SY_API_WEB"));
        log.debug("[SP_SY_API_WEB] out_ROW2 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW2")));
        resManager.setOutRow3(convertMap(map.get("out_ROW3"), "SP_SY_API_WEB"));
        log.debug("[SP_SY_API_WEB] out_ROW3 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW3")));

        return resManager;
    }

    /**
     * 관리웹사이트 전문 연결 API 처리 로직
     * @Method Name : spSyApiAdmWeb
     * @param reqManager
     * @param excelType
     * @return
     * @throws Exception
     */
    public Object spSyApiAdmWeb(ReqManager reqManager, boolean excelType) throws Exception {
        Map<String, Object> map = new HashMap<String, Object>();

        /** 요청타입 */
        map.put("REQ_TYPE", reqManager.getReqType());
        /** 요청전문 */
        map.put("in_REQ_NUM", reqManager.getReqNum());
        /** 언어구분 (MN_SY_CODE.SY_CODE = 'LT', 0001:한국어...) */
        map.put("in_LANGUAGE", reqManager.getLanguage());
        /** 요청한 단말구분 */
        map.put("in_DEVICE", reqManager.getDevice());
        /** 실제사용 정보 출력 Y/N (제휴본사, 총판, 가맹점 정보 등, API 서버를 서비스/테스트로 별도 운영하여 구분하여 세팅) */
        map.put("in_SERVICE_YN", reqManager.getServiceYn());

        /** 전달VALUE (DELIMETER : │) */
        map.put("in_VALUE", reqManager.getValue());

        /** 사용자정보 (DELIMETER : │, 전달값정의: USER_ID│USER_NAME│USER_TYPE_CD│HD_CODE│BR_CODE│ST_CODE│MULTI_TYPE│) */
        map.put("in_INFO", reqManager.getInfo());

        log.debug("[CALL SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+") JSON : "+CommonUtils.jsonStringFromObject(map));

        //관리웹사이트 전문 연결 API 호출
        managerMapper.spSyApiAdmWeb(map);

        log.info("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE : "+map.get("out_CODE")+", out_MSG : "+map.get("out_MSG"));

        if(map.get("out_CODE") == null){
            throw new Exception("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE is Null : "+map.get("out_CODE"));

        }else if(!NumberUtils.isDigits( String.valueOf(map.get("out_CODE")).trim() )){

            throw new Exception("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"(REQ_TYPE : "+reqManager.getReqType()+", REQ_NUM : "+reqManager.getReqNum()+")] out_CODE is Not Number : "+map.get("out_CODE"));

        }

        if(excelType){ //엑셀다운로드 요청일때

            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_CODE : "+CommonUtils.jsonStringFromObject(map.get("out_CODE")));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_MSG : "+CommonUtils.jsonStringFromObject(map.get("out_MSG")));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_VALUE : "+CommonUtils.jsonStringFromObject(map.get("out_VALUE")));
            //log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW1 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW1")));
            //log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW2 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW2")));
            //log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW3 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW3")));

            return map;

        }else{ //일반 전문 요청일때

            ResManager resManager = new ResManager();

            resManager.setOutCode(String.valueOf(map.get("out_CODE")));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_CODE : "+CommonUtils.jsonStringFromObject(map.get("out_CODE")));
            resManager.setOutMsg(String.valueOf(map.get("out_MSG")));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_MSG : "+CommonUtils.jsonStringFromObject(map.get("out_MSG")));
            resManager.setOutValue(String.valueOf(map.get("out_VALUE")));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_VALUE : "+CommonUtils.jsonStringFromObject(map.get("out_VALUE")));
            resManager.setOutRow1(convertMap(map.get("out_ROW1"), "SP_SY_API_ADM_WEB"));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW1 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW1")));
            resManager.setOutRow2(convertMap(map.get("out_ROW2"), "SP_SY_API_ADM_WEB"));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW2 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW2")));
            resManager.setOutRow3(convertMap(map.get("out_ROW3"), "SP_SY_API_ADM_WEB"));
            log.debug("[SP_SY_API_ADM_WEB"+CommonUtils.codeConv("REQ_TYPE", reqManager.getReqType()) +"] out_ROW3 : "+CommonUtils.jsonStringFromObject(map.get("out_ROW3")));

            return resManager;

        }

    }

    /**
     * 프로시저 결과값 컨버팅 ArrayList
     * @Method Name : convertMap
     * @param object
     * @return
     * @throws Exception
     */
    private ArrayList<LinkedHashMap<String, Object>> convertMap(Object object, String REQ_TYPE) throws Exception {
        if (object == null) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        ArrayList<LinkedHashMap<String, Object>> rows = (ArrayList<LinkedHashMap<String, Object>>) object;

        if (rows == null || rows.size() <= 0) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        if (null == rows.get(0) || "NULL".equals((rows.get(0).keySet().iterator().next()))) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        ArrayList<LinkedHashMap<String, Object>> arrayList = new ArrayList<LinkedHashMap<String, Object>>();

        int[] i = {0};
        for (LinkedHashMap<String, Object> row : rows) {
            LinkedHashMap<String, Object> resultMap = new LinkedHashMap<String, Object>();

            row.forEach((k,v)->{

                //log.debug("["+REQ_TYPE+"]["+i[0]+"] KEYNAME : "+ k);

                if(v != null){
                    //log.debug("["+REQ_TYPE+"]["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                    resultMap.put(k, String.valueOf(v));
                }
                else {
                    //log.debug("["+REQ_TYPE+"]["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                    resultMap.put(k, "");
                }

            });

            /*
            Iterator<String> iterator = row.keySet().iterator();

            String keyName;

            while (iterator.hasNext()) {
                keyName = iterator.next();

                log.debug("["+REQ_TYPE+"]["+i+"] KEYNAME : "+ keyName);

                if (row.get(keyName) != null){

                    log.debug("["+REQ_TYPE+"]["+i+"] KEY : "+ keyName+", VALUE : "+row.get(keyName));
                    resultMap.put(keyName, String.valueOf(row.get(keyName)));

                }else {
                    log.debug("["+REQ_TYPE+"]["+i+"] KEY : "+ keyName+", VALUE : ");
                    resultMap.put(keyName, "");
                }

            }
             */

            arrayList.add(resultMap);
            i[0]++;
        }

        return arrayList;
    }

}
