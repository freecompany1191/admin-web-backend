package com.o2osys.mng.packet.com;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

@Data
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ResCommon {

    /** 응답메세지 */
    @JsonProperty("error_msg")
    private String errorMsg;
}

