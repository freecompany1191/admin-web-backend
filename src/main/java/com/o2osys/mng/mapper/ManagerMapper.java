package com.o2osys.mng.mapper;

import java.util.Map;

import org.apache.ibatis.annotations.Mapper;

import com.o2osys.mng.config.Master;

@Mapper
@Master
public interface ManagerMapper {

    /**
     * 웹 관리자 로그인
     * @Method Name : spSyApiWeb
     * @param map
     */
    void spSyApiWeb(Map<String, Object> map);

    /**
     * 관리웹사이트 전문 연결
     * @Method Name : spSyApiAdmWeb
     * @param map
     */
    void spSyApiAdmWeb(Map<String, Object> map);

}
