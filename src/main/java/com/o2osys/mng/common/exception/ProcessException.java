package com.o2osys.mng.common.exception;

/**
   @FileName  : ProcessException.java
   @Description : 프로세스 예외처리
   @author      : KMS
   @since       : 2018. 1. 10.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 1. 10.    KMS            최초생성

 */
public class ProcessException extends RuntimeException {

    /** long */
    private static final long serialVersionUID = 3846875134800994124L;

    public ProcessException() {
        super("ProcessException");
    }

    public ProcessException(String message) {
        super(message);
    }

    public ProcessException(String message, Throwable cause) {
        super(message, cause);
    }

    public ProcessException(Throwable cause) {
        super(cause);
    }

    public ProcessException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
