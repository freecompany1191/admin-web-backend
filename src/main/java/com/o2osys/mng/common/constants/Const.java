package com.o2osys.mng.common.constants;

/**
   @FileName  : Const.java
   @Description : 관리자 웹 연동 상수
   @author      : KMS
   @since       : 2018. 2. 6.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 6.    KMS            최초생성

 */
public interface Const {

    /** 요청 전문 **/
    interface REQ {

        /** 전문 타입 */
        interface TYPE {

            /** 로그인API */
            final String LOGIN = "100";
            /** 통합API */
            final String API = "200";
            /** 통합API1 */
            final String API1 = "201";
            /** 통합API2 */
            final String API2 = "202";
            /** 통합API3 */
            final String API3 = "203";
            /** 통합API4 */
            final String API4 = "204";
        }

        /** 전문 타입 */
        interface TYPE_NM {

            /** 로그인API */
            final String LOGIN = "로그인API";
            /** 통합API */
            final String API = "통합API";
            /** 통합API */
            final String API1 = "통합API1";
            /** 통합API */
            final String API2 = "통합API2";
            /** 통합API */
            final String API3 = "통합API3";
            /** 통합API */
            final String API4 = "통합API4";

        }

    }


    /** 응답 코드 및 메세지 */
    interface RES {

        interface CODE {

            /** 정상(0000) */
            final String OK = "1";
            /** Request 파라미터 오류(0001) */
            final String REQ_ERROR="0001";
            /** 처리실패(0) */
            final String FAIL = "0";
            /** 기 처리상태(0003) */
            final String ALREADY_COMPLETE = "0003";
            /** 키값인증 오류(0004) */
            final String AUTH_ERROR="0004";
            /** 주문상태 불명(0005) */
            final String BAD_ORDER="0005";
            /** 접속 오류(0006) */
            final String NOT_CONNECT="0006";
            /** 시스템 오류(정의되지 않은 오류)(9999) */
            final String SYSTEM_ERROR="9999";
        }

        interface MSG {

            /** 정상(처리결과) - 0000 */
            final String OK = "정상 처리";
            final String NO_DATA = "조회된 데이터가 없습니다";
            /** Request 파라미터 오류(필수누락 등) - 0001 */
            final String REQ_ERROR="Request 파라미터 오류(필수누락 등)";
            /** 처리거절(처리 불가능 상황) - 0002 */
            final String FAIL = "처리거절(처리 불가능 상황)";
            /** 기 처리상태(이미 처리된 상태) - 0003 */
            final String ALREADY_COMPLETE = "기 처리상태(이미 처리된 상태)";
            /** 키값인증 오류 - 0004 */
            final String AUTH_ERROR="키값인증 오류";
            /** 주문상태 불명 - 0005 */
            final String BAD_ORDER="주문상태 불명";
            /** 접속 오류(0006) */
            final String NOT_CONNECT="접속 오류";
            /** 시스템 오류(정의되지 않은 오류) - 9999 */
            final String SYSTEM_ERROR="시스템 오류(정의되지 않은 오류)";
        }
    }


    /** BizTalk-Center API Response Result DATA */
    interface BIZ_TALK {

        /** BizTalk-Center API Response Result Const */
        interface RES {
            /** BizTalk-Center API Response Result Code */
            interface CODE {

                /** 요청 성공(200) */
                final String OK = "200";
                /** 권한 없음(403) */
                final String NO_AUTH="403";
                /** 파라미터 오류(405) */
                final String PARAM_ERROR="405";
                /** 템플릿 코드 중복(504) */
                final String DUPLE_CODE = "504";
                /** 템플릿 이름 중복(505) */
                final String DUPLE_NAME = "505";
                /** 템플릿 내용이 1000자 초과(506) */
                final String BYTE_ERROR= "506";
                /** 유효하지 않은 발신 프로필(507) */
                final String NO_SENDER = "507";
                /** 요청한 데이터가 없음(508) */
                final String NO_REQ_DATA = "508";
                /** 요청을 처리할 수 있는 상태가 아님(509)
                 *  (ex: 템플릿 검수 요청이 가능한 상태가 아닙니다.)
                 **/
                final String NO_STATUS = "509";
                /** BizTalk-Center 인증키 오류(900) */
                final String BIZ_AUTH_KEY_ERROR = "900";
                /** BizTalk-Center API 오류(901) */
                final String BIZ_API_ERROR="901";
                /** 네트워크 오류(991) */
                final String NOT_CONNECT="991";
                /** 시스템 오류(정의되지 않은 오류)(999) */
                final String SYSTEM_ERROR="999";
            }

            /** BizTalk-Center API Response Result Message */
            interface MSG {

                /** 요청 성공(200) */
                final String OK = "요청 성공";
                /** 권한 없음(403) */
                final String NO_AUTH="권한 없음";
                /** 파라미터 오류(405) */
                final String PARAM_ERROR="파라미터 오류";
                /** 템플릿 코드 중복(504) */
                final String DUPLE_CODE = "템플릿 코드 중복";
                /** 템플릿 이름 중복(505) */
                final String DUPLE_NAME = "템플릿 이름 중복";
                /** 템플릿 내용이 1000자 초과(506) */
                final String BYTE_ERROR= "템플릿 내용이 1000자 초과";
                /** 유효하지 않은 발신 프로필(507) */
                final String NO_SENDER = "유효하지 않은 발신 프로필";
                /** 요청한 데이터가 없음(508) */
                final String NO_REQ_DATA = "요청한 데이터가 없음";
                /** 요청을 처리할 수 있는 상태가 아님(509)
                 *  (ex: 템플릿 검수 요청이 가능한 상태가 아닙니다.)
                 **/
                final String NO_STATUS = "요청을 처리할 수 있는 상태가 아님(템플릿 검수 요청이 가능한 상태가 아님)";
                /** BizTalk-Center 인증키 오류(900) */
                final String BIZ_AUTH_KEY_ERROR = "BizTalk-Center 인증키 오류";
                /** BizTalk-Center API 오류(901) */
                final String BIZ_API_ERROR="BizTalk-Center API 오류";
                /** 네트워크 오류(991) */
                final String NOT_CONNECT="네트워크 오류";
                /** 시스템 오류(정의되지 않은 오류) - 999 */
                final String SYSTEM_ERROR="시스템 오류(정의되지 않은 오류)";
            }

            /** BizTalk-Center API 응답 데이터 */
            interface DATA {

                /** 템플릿 삭제 상태 */
                interface STATUS {

                    /** 템플릿 삭제 상태 코드 */
                    interface CODE {
                        /** 중단(S) */
                        final String STOP = "S";
                        /** 정상(A) */
                        final String OK = "A";
                        /** 대기(발송전)(R) */
                        final String WAIT = "R";
                    }

                    /** 템플릿 삭제 상태 메세지 */
                    interface MSG {
                        /** 중단(S) */
                        final String STOP = "중단";
                        /** 정상(A) */
                        final String OK = "정상";
                        /** 대기(발송전)(R) */
                        final String WAIT = "대기(발송전)";
                    }

                }

                /** 템플릿 검사 상태 */
                interface INSEPECT_STATUS {

                    /** 템플릿 검사 상태 코드 */
                    interface CODE {
                        /** 접수(REG) */
                        final String REG = "REG";
                        /** 등록(APL) */
                        final String APL = "APL";
                        /** 검수중(INS) */
                        final String INS = "INS";
                        /** 승인(COM) */
                        final String COM = "COM";
                        /** 반려(REJ) */
                        final String REJ = "REJ";
                    }

                    /** 템플릿 검사 상태 메세지 */
                    interface MSG {
                        /** 접수(REG) */
                        final String REG = "접수";
                        /** 등록(APL) */
                        final String APL = "등록";
                        /** 검수중(INS) */
                        final String INS = "검수중";
                        /** 승인(COM) */
                        final String COM = "승인";
                        /** 반려(REJ) */
                        final String REJ = "반려";
                    }

                }

                /** 템플릿 버튼 타입 */
                interface BUTTON_TYPE {

                    /** 템플릿 버튼 타입 코드 */
                    interface CODE {
                        /** 없음(default)(N) */
                        final String NONE = "N";
                        /** 배송조회(DS) */
                        final String DELIVERY = "DS";
                        /** 자유설정(C) */
                        final String FREE = "C";
                    }

                    /** 템플릿 버튼 타입 메세지 */
                    interface MSG {
                        /** 없음(default)(N) */
                        final String NONE = "없음(default)";
                        /** 배송조회(DS) */
                        final String DELIVERY = "배송조회";
                        /** 자유설정(C) */
                        final String FREE = "자유설정";
                    }

                }

            }

        }

        /** 그룹 템플릿 코드 */
        interface TPLT{

            /** 템플릿 구분 */
            interface TYPE {

                /** 템플릿 구분 코드 */
                interface CODE {
                    /** 주문하기(1) */
                    final String ORDER = "1";
                    /** 평가하기(2) */
                    final String EVAL = "2";
                    /** 주문접수(3) */
                    final String RECEIPT = "3";
                    /** 출발안내(4) */
                    final String DEPART = "4";
                }

                /** 템플릿 구분 메세지 */
                interface MSG {
                    /** 주문하기(1) */
                    final String ORDER = "주문하기";
                    /** 평가하기(2) */
                    final String EVAL = "평가하기";
                    /** 주문접수(3) */
                    final String RECEIPT = "주문접수";
                    /** 출발안내(4) */
                    final String DEPART = "출발안내";
                }

            }

            /** 템플릿 버튼 타입 */
            interface BUTTON_TYPE {

                /** 템플릿 버튼 타입 코드 */
                interface CODE {
                    /** 없음(default)(N) */
                    final String NONE = "N";
                    /** 배송조회(DS) */
                    final String DELIVERY = "DS";
                    /** 자유설정(C) */
                    final String FREE = "C";
                }

                /** 템플릿 버튼 타입 메세지 */
                interface MSG {
                    /** 없음(default)(N) */
                    final String NONE = "없음(default)";
                    /** 배송조회(DS) */
                    final String DELIVERY = "배송조회";
                    /** 자유설정(C) */
                    final String FREE = "자유설정";
                }

            }

            /** 템플릿 서비스 구분 */
            interface SERVICE_TYPE {

                /** 템플릿 서비스 구분 코드 */
                interface CODE {
                    /** 만나(1) */
                    final String MANNA = "1";
                    /** BBQ(2) */
                    final String BBQ = "2";
                }

                /** 템플릿 서비스 구분 메세지 */
                interface MSG {
                    /** 만나(1) */
                    final String MANNA = "만나";
                    /** BBQ(2) */
                    final String BBQ = "BBQ";
                }

            }

            /** 템플릿 센더 키 구분 */
            interface SENDER_KEY_TYPE {

                /** 템플릿 센더 키 구분 코드 */
                interface CODE {
                    /** 그룹(G) */
                    final String GROUP = "G";
                    /** 개별(S) */
                    final String SINGLE = "S";
                }

                /** 템플릿 센더 키 구분 메세지 */
                interface MSG {
                    /** 그룹(G) */
                    final String GROUP = "그룹";
                    /** 개별(S) */
                    final String SINGLE = "개별";
                }

            }

            /** 템플릿 데이터 상태 */
            interface DATA_STATUS {

                /** 템플릿 승인 상태 코드 */
                interface CODE {
                    /** 삭제(0) */
                    final String NO_DATA = "0";
                    /** 정상(1) */
                    final String OK = "1";
                }

                /** 템플릿 승인 상태 메세지 */
                interface MSG {
                    /** 삭제(0) */
                    final String NO_DATA = "삭제";
                    /** 정상(1) */
                    final String OK = "정상";
                }

            }

            /** 템플릿 승인 상태 */
            interface APPR_STATUS {

                /** 템플릿 승인 상태 코드 */
                interface CODE {
                    /** 승인요청(0) */
                    final String REQ = "0";
                    /** 승인완료(1) */
                    final String COMPLETE = "1";
                    /** 반려(2) */
                    final String REJECT = "2";
                    /** 승인요청실패(3) */
                    final String REQ_FAIL = "3";
                }

                /** 템플릿 승인 상태 메세지 */
                interface MSG {
                    /** 승인요청(0) */
                    final String REQ = "승인요청";
                    /** 승인완료(1) */
                    final String COMPLETE = "승인완료";
                    /** 반려(2) */
                    final String REJECT = "반려";
                    /** 승인요청실패(3) */
                    final String REQ_FAIL = "승인요청실패";
                }

            }

        }

    }



}
