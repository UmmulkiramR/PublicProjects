package com.cdplatform.document.utils;

import org.apache.commons.lang.StringUtils;

public class DocumentUtil {

    public static Long retreiveDocId(String docId) {
        if(StringUtils.isNotBlank(docId) && StringUtils.contains(docId,'.')){
            String value=docId.substring(0,docId.lastIndexOf('.'));
            return Long.valueOf(value);
        }
        return 0L;
    }
}
