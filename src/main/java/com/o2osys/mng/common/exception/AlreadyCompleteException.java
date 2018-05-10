package com.o2osys.mng.common.exception;

/**
   @FileName  : AlreadyCompleteException.java
   @Description : 이미 처리된 상태
   @author      : KMS
   @since       : 2018. 1. 10.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 10.    KMS            최초생성

 */
public class AlreadyCompleteException extends RuntimeException {

    /** long */
    private static final long serialVersionUID = 3846875134800994124L;

    public AlreadyCompleteException() {
        super("AlreadyCompleteException");
    }

    public AlreadyCompleteException(String message) {
        super(message);
    }

    public AlreadyCompleteException(String message, Throwable cause) {
        super(message, cause);
    }

    public AlreadyCompleteException(Throwable cause) {
        super(cause);
    }

    public AlreadyCompleteException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
