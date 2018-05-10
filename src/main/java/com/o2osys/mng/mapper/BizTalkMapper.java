package com.o2osys.mng.mapper;

import org.apache.ibatis.annotations.Mapper;

import com.o2osys.mng.config.Master;
import com.o2osys.mng.entity.GrpTpltVO;

@Mapper
@Master
public interface BizTalkMapper {

    /**
     * 그룹 템플릿 비즈톡 응답 데이터 업데이트
     * @Method Name : updateGrpTpltStatus
     * @param map
     */
    int updateGrpTpltStatus(GrpTpltVO grpTpltVO);

    /**
     * 그룹 템플릿 조회
     * @Method Name : selectGrpTplt
     * @param grpTpltVO
     */
    GrpTpltVO selectGrpTplt(GrpTpltVO grpTpltVO);

}
