package com.o2osys.mng.packet.mng;

import java.util.ArrayList;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.o2osys.mng.entity.MnGrSt;

import lombok.Data;

@Data
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ResMnGrStList {

    private int pageno;

    private int pagesize;

    private int totalcount;

    @JsonProperty("mngrsts")
    private ArrayList<MnGrSt> mnGrStList;
}
