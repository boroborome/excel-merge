package com.balance.excel.merge.util;

import com.balance.excel.merge.exception.MessageRecorderException;
import lombok.NoArgsConstructor;

import java.text.MessageFormat;
import java.util.*;

@NoArgsConstructor
public class MessageRecorder {
    private List<String> errors = new ArrayList<>();
    private List<String> warnings = new ArrayList<>();
    private MessageFormatter formatter;
    private boolean ignoreError;
    private boolean ignoreWarning;

    private final static String KeyError = "error";
    private final static String KeyWarning = "warning";

    private Map<String, Set<String>> errorKeys = MapBuilder.<String, Set<String>>of(KeyError, new HashSet<>())
            .and(KeyWarning, new HashSet<>())
            .build();

    public MessageRecorder(boolean ignoreWarning) {
        this.ignoreWarning = ignoreWarning;
    }

    public MessageRecorder(boolean ignoreWarning, boolean ignoreError) {
        this.ignoreError = ignoreError;
        this.ignoreWarning = ignoreError ? true : ignoreWarning;
    }

    public MessageFormatter getFormatter() {
        return formatter;
    }

    public void setFormatter(MessageFormatter formatter) {
        this.formatter = formatter;
    }

    public void appendDistinctError(String messageKey, String format, Object... parameters) {
        appendMessage(errors, KeyError, messageKey, format, parameters);
    }
    public void appendError(String format, Object... parameters) {
        appendMessage(errors, KeyError, null, format, parameters);
    }

    public void appendDistinctWarning(String messageKey, String format, Object... parameters) {
        appendMessage(warnings, KeyWarning, messageKey, format,  parameters);
    }
    public void appendWarning(String format, Object... parameters) {
        appendMessage(warnings, KeyWarning,null, format, parameters);
    }

    protected void appendMessage(List<String> messages, String messageType, String errorKey, String format, Object... parameters) {
        if (errorKey != null) {
            Set<String> existMessage = errorKeys.get(messageType);
            if (existMessage.contains(errorKey)) {
                return;
            }
            existMessage.add(errorKey);
        }
        String message = MessageFormat.format(format, parameters);
        if (formatter != null) {
            message = formatter.format(message);
        }
        messages.add(message);
    }

    public List<String> getErrors() {
        return errors;
    }

    public List<String> getWarnings() {
        return warnings;
    }

    public boolean isSuccess() {
        return (ignoreError || errors.isEmpty())
                && (ignoreWarning || warnings.isEmpty());
    }

    public boolean containsError(String error) {
        if (errors.isEmpty()) {
            return false;
        }
        for (String message : errors) {
            if (Objects.equals(error, message)) {
                return true;
            }
        }
        return false;
    }

    public void distinctMessage() {
        distinctMessage(errors);
        distinctMessage(warnings);
    }

    private void distinctMessage(List<String> messages) {
        Set<String> set = new HashSet<>();
        Iterator<String> it = messages.iterator();
        while (it.hasNext()) {
            String message = it.next();
            if (set.contains(message)) {
                it.remove();
            } else {
                set.add(message);
            }
        }
    }

    public static MessageRecorder ignoreWarnings() {
        MessageRecorder messageRecorder = new MessageRecorder();
        messageRecorder.ignoreError = false;
        messageRecorder.ignoreWarning = true;
        return messageRecorder;
    }

    public static MessageRecorder ignoreErrors() {
        MessageRecorder messageRecorder = new MessageRecorder();
        messageRecorder.ignoreError = true;
        messageRecorder.ignoreWarning = true;
        return messageRecorder;
    }

    public static MessageRecorder exceptionErrors() {
        return new ExceptionMessageRecorder();
    }

    private static class ExceptionMessageRecorder extends MessageRecorder {
        @Override
        protected void appendMessage(List<String> messages, String messageType, String errorKey, String format, Object... parameters) {
            super.appendMessage(messages, messageType, errorKey, format, parameters);
            if (!this.isSuccess()) {
                throw new MessageRecorderException(this.getErrors().get(0));
            }
        }
    }

    public void reset() {
        errors.clear();
        warnings.clear();
    }
    public interface MessageFormatter {
        String format(String message);
    }

}
