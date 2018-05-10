package com.o2osys.mng.common.constants;

/**
   @FileName  : ConstFile.java
   @Description : 파일 서비스 상수
   @author      : KMS
   @since       : 2018. 2. 6.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 6.    KMS            최초생성

 */
public interface ConstFile {

    /** 문서 파일 설정 */
    interface DOCUMENT {
        /** 문서 파일 경로 */
        interface PATH {

        }
    }

    /** 기타 파일 설정 */
    interface OTHER {

        /** 기타 파일 경로 */
        interface PATH {

        }
    }

    /** 이미지 파일 설정 */
    interface IMAGE {

        /** 이미지 파일 경로 */
        interface PATH {

            /** 썸네일 이미지 파일 경로 */
            final String THUMBNAIL = "Thumbnail/";

            /** 가맹점 대표 이미지 파일 경로 */
            final String STORE = "store/";
            /** 가맹점 인트로 이미지 파일 경로 */
            final String STORE_INTRO = "store_intro/";
            /** 가맹점 상품 이미지 파일 경로 */
            final String STORE_GOODS = "store_goods/";
            /** 가맹점 이벤트 이미지 파일 경로 */
            final String STORE_EVENT = "store_event/";

            /** 리뷰 이미지 파일 경로 */
            final String REVIEW = "review/";

            /** 사업자 등록증 이미지 파일 경로 */
            final String REGISTRATION = "registration/";

            /** 제휴사공지 이미지 파일 경로 */
            final String BBS_B2B = "bbs/B2B/";
            /** 플랫폼공지 이미지 파일 경로 */
            final String BBS_SYS = "bbs/SYS/";
            /** 고객공지 이미지 파일 경로 */
            final String BBS_CU = "bbs/CU/";
            /** 일대일 이미지 파일 경로 */
            final String BBS_CU_ONEANDONE = "bbs/CU/ONEANDONE/";

            /** 배송원 이미지 파일 경로 */
            final String WORKER = "worker/";
            /** 배송원 근로계약서 */
            final String WORKER_CONTRACT = "contract/worker/";
            /** 배송원 법정동의서 */
            final String WORKER_CONSENT = "consent/worker/";

            /** 종업원 이미지 */
            final String EMPLOYEE = "employee/";
            /** 종업원 근로계약서 */
            final String EMPLOYEE_CONTRACT = "contract/employee/";
            /** 종업원 법정동의서 */
            final String EMPLOYEE_CONSENT = "consent/employee/";

        }

        /** 이미지 사이즈 */
        interface SIZE {

            /** 썸네일 이미지 사이즈 */
            interface THUMBNAIL {
                /** 넓이 */
                final String WIDTH = "239";
                /** 높이 */
                final String HEIGHT = "239";
            }

            /** 가맹점 대표 이미지 사이즈 */
            interface STORE {
                /** 넓이 */
                final String WIDTH = "239";
                /** 높이 */
                final String HEIGHT = "239";
            }

            /** 가맹점 인트로 이미지 사이즈 */
            interface STORE_INTRO {
                /** 넓이 */
                final String WIDTH = "1080";
                /** 높이 */
                final String HEIGHT = "1920";
            }

            /** 가맹점 상품 이미지 사이즈 */
            interface STORE_GOODS {
                /** 넓이 */
                final String WIDTH = "902";
                /** 높이 */
                final String HEIGHT = "677";
            }

            /** 리뷰 이미지 사이즈 */
            interface REVIEW {
                /** 넓이 */
                final String WIDTH = "984";
                /** 높이 */
                final String HEIGHT = "984";
            }

            /** 사업자 등록증 이미지 사이즈 */
            interface REGISTRATION {
                /** 넓이 */
                final String WIDTH = "984";
                /** 높이 */
                final String HEIGHT = "984";
            }

            /** 일대일 이미지 사이즈 */
            interface BBS_CU_ONEANDONE {
                /** 넓이 */
                final String WIDTH = "984";
                /** 높이 */
                final String HEIGHT = "984";
            }

            /** 배송원 이미지 사이즈 */
            interface WORKER {
                /** 넓이 */
                final String WIDTH = "239";
                /** 높이 */
                final String HEIGHT = "239";
            }

            /** 종업원 이미지 */
            interface EMPLOYEE {
                /** 넓이 */
                final String WIDTH = "239";
                /** 높이 */
                final String HEIGHT = "239";
            }

        }

    }

}
