package com.o2osys.mng.mapper;

import java.util.ArrayList;
import java.util.Map;

import org.apache.ibatis.annotations.Mapper;

import com.o2osys.mng.config.Master;
import com.o2osys.mng.entity.MnGrSt;

@Mapper
@Master
public interface MngMapper {

    /**
     * 가맹점 리스트
     *
     * @Method Name : mnGrStList
     * @param map
     * @return
     */
    ArrayList<MnGrSt> mnGrStList(Map<String, Object> map);

    /**
     * 가맹점 리스트 갯수
     * @Method Name : mnGrStListSize
     * @param map
     * @return
     */
    int mnGrStListSize();

}
