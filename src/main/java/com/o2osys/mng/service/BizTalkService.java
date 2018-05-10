package com.o2osys.mng.service;

import java.util.ArrayList;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import com.o2osys.mng.common.constants.Const;
import com.o2osys.mng.common.constants.Const.RES;
import com.o2osys.mng.common.exception.AlreadyCompleteException;
import com.o2osys.mng.common.exception.ProcessException;
import com.o2osys.mng.common.exception.RequestException;
import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.entity.GrpTpltVO;
import com.o2osys.mng.mapper.BizTalkMapper;
import com.o2osys.mng.packet.manager.ResManager;
import com.o2osys.mng.packet.manager.bizTalk.ReqBizTalk;
import com.o2osys.mng.packet.manager.bizTalk.ReqTemplateSend;
import com.o2osys.mng.packet.manager.bizTalk.ResBizTalk;;

@Service("BizTalkService")
public class BizTalkService {
    // 로그
    private final Logger log = LoggerFactory.getLogger(BizTalkService.class);
    private final String TAG = BizTalkService.class.getSimpleName();

    @Autowired
    BizTalkMapper bizTalkMapper;

    @Autowired
    BizTalkSender bizTalkSender;

    @Autowired
    CommonService commonService;

    /**
     * 비즈톡 그룹 템플릿 관리 기능
     * @Method Name : sendTemplate
     * @param reqTemplateSend
     * @return
     * @throws Exception
     */
    public ResManager sendTemplate(ReqTemplateSend reqTemplateSend) throws Exception {

        //응답객체 선언
        ResManager resManager = new ResManager();
        //비즈톡 전송객체 선언
        ReqBizTalk reqBizTalk = new ReqBizTalk();
        //그룹 템플릿 객체
        GrpTpltVO grpTpltVO = new GrpTpltVO();
        //비즈톡 응답 객체
        ResBizTalk resBizTalk = new ResBizTalk();

        try {

            if(StringUtils.isEmpty(reqTemplateSend)){
                throw new RequestException(RES.MSG.REQ_ERROR+":: NULL");
            }

            if(StringUtils.isEmpty(reqTemplateSend.getTpltCode())){
                throw new RequestException(RES.MSG.REQ_ERROR+":: TPLT_CODE");
            }

            if(StringUtils.isEmpty(reqTemplateSend.getTranType())){
                throw new RequestException(RES.MSG.REQ_ERROR+":: TRAN_TYPE");
            }

            //템플릿 코드 셋팅
            reqBizTalk.setTemplateCode(reqTemplateSend.getTpltCode());
            //전송 타입
            String tranType = reqTemplateSend.getTranType();

            switch(tranType){
                case "CREATE" :
                    //비즈톡 템플릿 등록 전송 처리
                    resBizTalk = sendCreate(reqTemplateSend) ;
                    break;
                case "SEARCH" :
                    //비즈톡 템플릿 조회 전송 처리
                    resBizTalk = bizTalkSender.searchtemplate(reqBizTalk);
                    break;
            }

            //그룹 템플릿 코드 셋팅
            grpTpltVO.setTpltCode(reqTemplateSend.getTpltCode());
            //응답 유저 ID셋팅
            grpTpltVO.setUserId(reqTemplateSend.getUserId());
            //응답 발신프로필 키
            //grpTpltVO.setSenderKey(SENDER_KEY);

            //응답 에러 메세지
            grpTpltVO.setErrorMsg(!StringUtils.isEmpty(resBizTalk.getMessage()) ? resBizTalk.getMessage() : "");

            //응답데이터가 있을때만 승인상태 변환처리
            if(resBizTalk.getData() != null){

                //검사상태(REG:접수, APL:등록, INS:검수중, COM:승인, REJ: 반려)에 따른 승인상태 변환
                String INSPECTION_STATUS = resBizTalk.getData().getInspectionStatus();

                if(StringUtils.isEmpty(INSPECTION_STATUS)){
                    //등록 요청시에 승인요청이 실패하면 실패처리
                    if("CREATE".equals(tranType))
                        grpTpltVO.setApprStatus(Const.BIZ_TALK.TPLT.APPR_STATUS.CODE.REQ_FAIL); //승인요청실패
                }

                switch(INSPECTION_STATUS){
                    case Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.CODE.REG : //접수
                        log.info("BIZ_TALK RESPONSE INSPECTION_STATUS :: REG - 접수");
                        grpTpltVO.setErrorMsg(Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.MSG.REG);
                        //승인요청상태 유지
                        break;
                    case Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.CODE.APL : //등록
                        log.info("BIZ_TALK RESPONSE INSPECTION_STATUS :: APL - 등록");
                        grpTpltVO.setErrorMsg(Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.MSG.APL);
                        //승인요청상태 유지
                        break;
                    case Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.CODE.INS : //검수중
                        log.info("BIZ_TALK RESPONSE INSPECTION_STATUS :: INS - 검수중");
                        grpTpltVO.setErrorMsg(Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.MSG.INS);
                        //승인요청상태 유지
                        break;
                    case Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.CODE.COM : //승인
                        log.info("BIZ_TALK RESPONSE INSPECTION_STATUS :: COM - 승인");
                        grpTpltVO.setApprStatus(Const.BIZ_TALK.TPLT.APPR_STATUS.CODE.COMPLETE); //승인완료
                        break;
                    case Const.BIZ_TALK.RES.DATA.INSEPECT_STATUS.CODE.REJ : //반려
                        log.info("BIZ_TALK RESPONSE INSPECTION_STATUS :: REJ - 반려");
                        grpTpltVO.setApprStatus(Const.BIZ_TALK.TPLT.APPR_STATUS.CODE.REJECT); //반려
                        grpTpltVO.setErrorMsg(resBizTalk.getData().getComments()); //반려 메세지
                        break;
                }

            }
            else{
                switch(tranType){
                    case "CREATE" :
                        //등록 요청시에 승인요청이 실패하면 실패처리
                        grpTpltVO.setApprStatus(Const.BIZ_TALK.TPLT.APPR_STATUS.CODE.REQ_FAIL); //승인요청실패
                        break;
                    case "SEARCH" :

                        switch(resBizTalk.getCode()){
                            //요청한 데이터가 없음(508)
                            case Const.BIZ_TALK.RES.CODE.NO_REQ_DATA :
                                //조회 요청시에 응답 코드가 요청한 데이터가 없음(508) 실패처리
                                grpTpltVO.setApprStatus(Const.BIZ_TALK.TPLT.APPR_STATUS.CODE.REQ_FAIL); //승인요청실패
                                break;
                        }
                        break;
                }
            }
            //그룹 템플릿 비즈톡 응답 데이터 업데이트
            bizTalkMapper.updateGrpTpltStatus(grpTpltVO);

            //그룹 템플릿 조회
            GrpTpltVO resGrpTplt = bizTalkMapper.selectGrpTplt(grpTpltVO);

            //그룹 템플릿 조회 데이터 리턴 리스트 셋팅
            ArrayList<GrpTpltVO> resultList = new ArrayList<GrpTpltVO>();
            resultList.add(resGrpTplt);

            resManager.setOutCode(RES.CODE.OK);
            //resManager.setOutMsg(CommonUtils.codeConv("BIZ_TALK_RES", resBizTalk.getCode()));
            resManager.setOutMsg(grpTpltVO.getErrorMsg());
            resManager.setOutRow1(resultList);
            resManager.setOutRow2(new ArrayList());
            resManager.setOutRow3(new ArrayList());

            return resManager;

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
     * 비즈톡 등록 처리 후 응답 리턴
     * @Method Name : sendCreate
     * @param reqTemplateSend
     * @return
     * @throws Exception
     */
    private ResBizTalk sendCreate(ReqTemplateSend reqTemplateSend) throws Exception {

        //비즈톡 전송객체 선언
        ReqBizTalk reqBizTalk = new ReqBizTalk();

        if(StringUtils.isEmpty(reqTemplateSend.getTpltTitle())){
            throw new RequestException(RES.MSG.REQ_ERROR+":: TPLT_TITLE");
        }

        if(StringUtils.isEmpty(reqTemplateSend.getTpltContent())){
            throw new RequestException(RES.MSG.REQ_ERROR+":: TPLT_CONTENT");
        }

        //템플릿 코드 셋팅
        reqBizTalk.setTemplateCode(reqTemplateSend.getTpltCode());
        //템플릿 이름 셋팅
        reqBizTalk.setTemplateName(reqTemplateSend.getTpltTitle());
        //템플릿 내용 셋팅
        reqBizTalk.setTemplateContent(reqTemplateSend.getTpltContent());
        //템플릿 버튼명 셋팅
        reqBizTalk.setButtonName(reqTemplateSend.getBtnInfo());

        //버튼명 여부에 따른 타입 설정
        if(StringUtils.isEmpty(reqTemplateSend.getBtnInfo())){
            reqBizTalk.setButtonType(Const.BIZ_TALK.TPLT.BUTTON_TYPE.CODE.NONE); //없음(default)
        }else {
            reqBizTalk.setButtonType(Const.BIZ_TALK.TPLT.BUTTON_TYPE.CODE.FREE); //자유설정
        }

        //비즈톡 템플릿 등록 요청 전송 기능 호출
        ResBizTalk resBizTalk = bizTalkSender.createTemplate(reqBizTalk);

        return resBizTalk;

    }

}
