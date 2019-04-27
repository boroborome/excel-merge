package com.balance.excel.merge.exception;

public class MessageRecorderException extends RuntimeException {

    public MessageRecorderException(String message) {
        super(message);
    }

    public MessageRecorderException(String message, Throwable cause) {
        super(message, cause);
    }
}