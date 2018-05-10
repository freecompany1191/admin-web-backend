package com.o2osys.mng.common.exception;

/**
   @FileName  : BizTalkException.java
   @Description : 비즈톡 에러 예외 처리
   @author      : KMS
   @since       : 2018. 2. 1.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 1.    KMS            최초생성

 */
public class BizTalkException extends RuntimeException {

    /** long */
    private static final long serialVersionUID = -2933241522421707626L;

    public BizTalkException() {
        super("ProcessException");
    }

    public BizTalkException(String message) {
        super(message);
    }

    public BizTalkException(String message, Throwable cause) {
        super(message, cause);
    }

    public BizTalkException(Throwable cause) {
        super(cause);
    }

    public BizTalkException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
