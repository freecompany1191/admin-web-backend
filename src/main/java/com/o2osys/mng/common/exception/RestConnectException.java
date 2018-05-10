package com.o2osys.mng.common.exception;

/**
   @FileName  : RestConnectException.java
   @Description : REST 전송 에러 예외처리
   @author      : KMS
   @since       : 2018. 2. 1.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2018. 2. 1.    KMS            최초생성

 */
public class RestConnectException extends RuntimeException {

    /** long */
    private static final long serialVersionUID = 5585635455823434687L;

    public RestConnectException() {
        super("RestConnectException");
    }

    public RestConnectException(String message) {
        super(message);
    }

    public RestConnectException(String message, Throwable cause) {
        super(message, cause);
    }

    public RestConnectException(Throwable cause) {
        super(cause);
    }

    public RestConnectException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
