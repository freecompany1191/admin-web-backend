package com.o2osys.mng.controller;

import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.o2osys.mng.common.constants.Const.RES;
import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.packet.manager.ReqManager;
import com.o2osys.mng.packet.manager.ResManager;
import com.o2osys.mng.packet.manager.bizTalk.ReqTemplateSend;
import com.o2osys.mng.packet.manager.file.ReqFileManager;
import com.o2osys.mng.service.BizTalkService;
import com.o2osys.mng.service.ExcelViewService;
import com.o2osys.mng.service.ManagerService;

import io.swagger.annotations.ApiOperation;

/**
   @FileName  : ManagerController.java
   @Description : 웹 관리 컨트롤러
   @author      : KMS
   @since       : 2018. 1. 11.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 11.    KMS            최초생성

 */
@RestController
@RequestMapping(value = "/v1/manager")
public class ManagerController {
    // 로그
    private final Logger log = LoggerFactory.getLogger(ManagerController.class);
    private final String TAG = ManagerController.class.getSimpleName();

    private ObjectMapper mObjectMapper = new ObjectMapper();

    /** 공통서비스 */
    @Autowired
    private CommonService commonService;

    @Autowired
    ManagerService managerService;

    @Autowired
    BizTalkService bizTalkService;

    @Autowired
    ExcelViewService excelViewService;


    @ApiOperation(value = "관리자 웹 사이트 전문 연결 API", notes = "로그인 처리는 REQ_TYPE 을 100 으로 설정 API는 200로 설정",
            response = ResManager.class)
    @PostMapping(path = "/api", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public @ResponseBody ResManager callApis(
            @RequestBody ReqManager reqManager
            ) throws Exception {

        ResManager resManager = new ResManager();

        try {

            log.info("[API] REQ_JSON : "+CommonUtils.jsonStringFromObject(reqManager));

            resManager = managerService.callApi(reqManager, null, null, null);

            log.info("[API] RES_JSON : "+CommonUtils.jsonStringFromObject(resManager));

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

        return resManager;

    }


    @ApiOperation(value = "관리자 웹 사이트 전문&파일 업로드 겸용 연결 API", notes = "로그인 처리는 REQ_TYPE 을 100 으로 설정 API는 200로 설정",
            response = ResManager.class)
    @PostMapping(path = "/fileApi", consumes = { "multipart/mixed", "multipart/form-data" } , produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public @ResponseBody ResManager callApis(
            @RequestHeader HttpHeaders headers,
            @RequestPart(name="param",required = false) String param,
            @RequestPart(name="file",required = false) MultipartFile file
            ) throws Exception {

        log.debug("reqData : "+param);
        log.debug("reqFile : "+file);
        ReqFileManager reqFileManager = new ReqFileManager();
        ResManager resManager = new ResManager();

        try {

            if(param != null){
                param = URLDecoder.decode(param,"UTF-8");
                reqFileManager = new ObjectMapper().readValue(param, ReqFileManager.class);
                log.info("[API] REQ_JSON : "+CommonUtils.jsonStringFromObject(reqFileManager));

                resManager = managerService.callApi(null, reqFileManager, file, "fileApi");

            }else {
                /** Request 파라미터 오류(필수누락 등) - 0001 */
                log.error(RES.MSG.REQ_ERROR);
                resManager.setOutCode(RES.CODE.FAIL);
                resManager.setOutMsg(RES.MSG.REQ_ERROR);
                resManager.setOutRow1(new ArrayList());
                resManager.setOutRow2(new ArrayList());
                resManager.setOutRow3(new ArrayList());
            }

            log.info("[API] RES_JSON : "+CommonUtils.jsonStringFromObject(resManager));
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

        return resManager;

    }

    @ApiOperation(value = "관리자 웹 사이트 그룹 템플릿 API", notes = "비즈톡 STATUS 처리",
            response = ResManager.class)
    @PostMapping(path = "/template/bizTalkSender", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public @ResponseBody ResManager bizTalkSender (
            @RequestBody ReqTemplateSend reqTemplateSend
            ) throws Exception {

        ResManager resManager = new ResManager();
        try {

            log.info("[BIZTALK_API] REQ_JSON : "+CommonUtils.jsonStringFromObject(reqTemplateSend));

            resManager = bizTalkService.sendTemplate(reqTemplateSend);

            log.info("[BIZTALK_API] RES_JSON : "+CommonUtils.jsonStringFromObject(resManager));

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

        return resManager;

    }

    @ApiOperation(value = "관리자 웹 사이트 전문 연결 엑셀 API", notes = "관리자 웹 사이트 전문 연결 엑셀 API",
            response = ResManager.class)
    @PostMapping(path = "/excelDownApi")
    public ModelAndView callExcelApis(
            @RequestBody ReqManager reqManager
            ) throws Exception {

        Map<String, Object> insMap = new HashMap<String, Object>();

        try {

            log.info("[API] REQ_JSON : "+CommonUtils.jsonStringFromObject(reqManager));

            insMap = excelViewService.getExcel(reqManager);

            //log.info("[API] RES_JSON : "+CommonUtils.jsonStringFromObject(insMap));

        } catch (Exception e) {
            /** 시스템 오류(정의되지 않은 오류) - 9999 */
            commonService.errorLog(TAG, e);
            return new ModelAndView("excelXlsxView", insMap);
        }

        return new ModelAndView("excelXlsxView", insMap);

    }


}