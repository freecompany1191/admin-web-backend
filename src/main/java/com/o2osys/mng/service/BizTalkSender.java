package com.o2osys.mng.service;

import java.net.ConnectException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.web.client.ResourceAccessException;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.o2osys.mng.common.constants.Const;
import com.o2osys.mng.common.constants.Define;
import com.o2osys.mng.common.exception.RestConnectException;
import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.common.service.LogService;
import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.mapper.BizTalkMapper;
import com.o2osys.mng.packet.manager.bizTalk.ReqBizTalk;
import com.o2osys.mng.packet.manager.bizTalk.ResBizTalk;

@Component
public class BizTalkSender {
    // 로그
    private final Logger log = LoggerFactory.getLogger(BizTalkSender.class);
    private final String TAG = BizTalkSender.class.getSimpleName();

    @Autowired
    BizTalkMapper bizTalkMapper;

    @Autowired
    LogService logService;

    @Autowired
    CommonService commonService;

    /** 비즈톡 템플릿 등록 요청 POST URL */
    @Value("${biztalk.sender.create.rest.url}")
    String SENDER_CREATE_URL;

    /** 비즈톡 템플릿 조회 요청 GET URL */
    @Value("${biztalk.sender.search.rest.url}")
    String SENDER_SEARCH_URL;

    /** 비즈톡 발신프로필 키 */
    @Value("${biztalk.sender.key}")
    String SENDER_KEY;

    /** biztalk-center에서 발급 받은 사이트아이디 */
    @Value("${biztalk.siteid}")
    String SITE_ID;

    /** biztalk-center에서 발급 받은 인증키 */
    @Value("${biztalk.auth.key}")
    String AUTH_KEY;


    /**
     * 카카오 비즈톡 템플릿 등록
     * @Method Name : bizTalkCreate
     * @param reqBizTalk
     * @return
     * @throws Exception
     */
    public ResBizTalk createTemplate(ReqBizTalk reqBizTalk) throws Exception {

        return RestTransfer(reqBizTalk, "CREATE");
    }

    /**
     * 카카오 비즈톡 템플릿 조회
     * @Method Name : bizTalkSearch
     * @param reqBizTalk
     * @return
     * @throws Exception
     */
    public ResBizTalk searchtemplate(ReqBizTalk reqBizTalk) throws Exception {

        return RestTransfer(reqBizTalk, "SEARCH");
    }

    /**
     * 카카오 비즈톡 전송 및 응답 처리
     * @Method Name : RestTransfer
     * @param reqBizTalk
     * @param transType
     * @return
     * @throws Exception
     */
    private ResBizTalk RestTransfer(ReqBizTalk reqBizTalk, String transType) throws Exception {

        HttpStatus resStatusCode = null;

        //REST 응답 객체
        ResponseEntity<String> resStr = null;
        //REST 응답 JSON 파싱 객체
        ResBizTalk resMsg;
        //리턴 객체 셋팅
        ResBizTalk res = new ResBizTalk();

        try {

            //요청 정보를 전송하고 응답 메세지를 받아온다
            resStr = getResMsg(reqBizTalk, transType);
            resMsg = new ObjectMapper().readValue(resStr.getBody(), ResBizTalk.class);
            resStatusCode = resStr.getStatusCode();

            //전송 성공일 때만 셋팅
            if (resStatusCode == HttpStatus.OK) {

                res.setCode(resMsg.getCode()); //응답 코드 셋팅
                res.setMessage(resMsg.getMessage()); //응답 메세지 셋팅
                res.setData(resMsg.getData()); //응답 데이터 셋팅

                if(res.getCode() != null){

                    //응답 코드가 성공이 아니면 에러 메세지 셋팅
                    if(!Const.BIZ_TALK.RES.CODE.OK.equals(res.getCode())){

                        log.error("["+transType+"] BIZ_TALK RESPONSE ERROR :: TEMPLATE_CODE = "+reqBizTalk.getTemplateCode()+", RES_CODE : "+res.getCode()+", ERROR_MSG = " + res.getMessage());
                        logService.error(Define.Log.TYPE, TAG, "["+transType+"] BIZ_TALK RESPONSE ERROR :: TEMPLATE_CODE = "+reqBizTalk.getTemplateCode()+", RES_CODE : "+res.getCode()+", ERROR_MSG = " + res.getMessage());

                    }

                }else throw new RestConnectException("["+transType+"] BIZ_TALK REST_ERROR :: TEMPLATE_CODE = "+reqBizTalk.getTemplateCode()+", ERROR_MSG = " + Const.BIZ_TALK.RES.MSG.NOT_CONNECT);

            }else throw new RestConnectException("["+transType+"] BIZ_TALK REST_CONNECT_ERROR :: TEMPLATE_CODE = "+reqBizTalk.getTemplateCode()+", ERROR_MSG = " + Const.BIZ_TALK.RES.MSG.NOT_CONNECT);

            //응답 JSON 로그
            log.info("["+transType+"] BIZ_TALK REST RES JSON :: "+CommonUtils.jsonStringFromObject(res));
            return res;

        } catch (RestConnectException | ResourceAccessException | ConnectException e) {
            commonService.errorLog(TAG, e);
            res.setCode(Const.BIZ_TALK.RES.CODE.NOT_CONNECT);
            res.setErrorMsg(Const.BIZ_TALK.RES.MSG.NOT_CONNECT);
            logService.error(Define.Log.TYPE, TAG, e.getMessage());
            return res;

        } catch (Exception e) {
            commonService.errorLog(TAG, e);
            res.setCode(Const.BIZ_TALK.RES.CODE.SYSTEM_ERROR);
            res.setErrorMsg(Const.BIZ_TALK.RES.MSG.SYSTEM_ERROR);
            logService.error(Define.Log.TYPE, TAG, "["+transType+"] BIZ_TALK REST_ERROR :: TEMPLATE_CODE = "+reqBizTalk.getTemplateCode()+", ERROR_MSG = " + e.getMessage());
            return res;
        }

    }

    /**
     * 전송 타입에 따라 요청 정보를 전송하고 응답 메세지를 받아온다
     * @Method Name : getResMsg
     * @param reqBizTalk
     * @param transType
     * @return
     * @throws Exception
     */
    private ResponseEntity<String> getResMsg(ReqBizTalk reqBizTalk, String transType) throws Exception{

        RestTemplate restTemplate = new RestTemplate();

        //헤더값 셋팅
        HttpHeaders headers = new HttpHeaders();
        //headers.setContentType(MediaType.APPLICATION_JSON_UTF8);
        headers.set("auth_key", AUTH_KEY); //biztalk-center에서 발급 받은 인증키 셋팅
        headers.set("siteid"  , SITE_ID ); //biztalk-center에서 발급 받은 사이트아이디 셋팅
        headers.setContentType(MediaType.APPLICATION_FORM_URLENCODED);  //application/x-www-form-urlencoded 방식으로 셋팅

        reqBizTalk.setSenderKey(SENDER_KEY); //비즈톡 발신프로필 키 셋팅
        reqBizTalk.setSenderKeyType(Const.BIZ_TALK.TPLT.SENDER_KEY_TYPE.CODE.GROUP); //비즈톡 발신프로필 키 타입 그룹 셋팅

        //전송 객체 셋팅
        HttpEntity<Object> entity = new HttpEntity<Object>(headers);
        log.info("["+transType+"]REST REQ ENTITY HEADERS JSON :: "+CommonUtils.jsonStringFromObject(entity.getHeaders()));




        //응답 객체 셋팅(비즈톡에서 스트링 형태로 리턴해줌)
        ResponseEntity<String> resStr = null;

        UriComponentsBuilder builder = null;

        switch(transType){
            case "CREATE" : //템플릿 등록 요청
                builder = UriComponentsBuilder.fromHttpUrl(SENDER_CREATE_URL)
                .queryParam("senderKey", reqBizTalk.getSenderKey())
                .queryParam("templateCode", reqBizTalk.getTemplateCode())
                .queryParam("templateName", reqBizTalk.getTemplateName())
                .queryParam("templateContent", reqBizTalk.getTemplateContent())
                .queryParam("senderKeyType", reqBizTalk.getSenderKeyType())
                .queryParam("buttonType", reqBizTalk.getButtonType());
                //요청 REST REQ QUERY 로그
                log.info("["+transType+"]REST REQ QUERY :: "+builder.build().encode().toUri().getQuery());

                resStr = restTemplate.exchange(builder.build().encode().toUri(), HttpMethod.POST, entity, String.class);
                break;
            case "SEARCH" : //템플릿 조회 요청
                builder = UriComponentsBuilder.fromHttpUrl(SENDER_SEARCH_URL)
                .queryParam("senderKey", reqBizTalk.getSenderKey())
                .queryParam("templateCode", reqBizTalk.getTemplateCode())
                .queryParam("senderKeyType", reqBizTalk.getSenderKeyType());
                //요청 REST REQ QUERY 로그
                log.info("["+transType+"]REST REQ QUERY :: "+builder.build().encode().toUri().getQuery());

                resStr = restTemplate.exchange(builder.build().encode().toUri(), HttpMethod.POST, entity, String.class);
                break;
        }


        //요청 REST REQ QUERY 로그
        //log.info("["+transType+"]BIZ_TALK REST REQ JSON :: "+CommonUtils.jsonStringFromObject(reqBizTalk));

        return resStr;

    }

}
