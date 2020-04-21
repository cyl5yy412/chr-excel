package com.chryl.service;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.chryl.mapper.LoanInfoMapper;
import com.chryl.po.LoanInfo;
import org.springframework.stereotype.Service;

/**
 * Created by Chr.yl on 2020/4/21.
 *
 * @author Chr.yl
 */
@Service
public class LoanInfoService extends ServiceImpl<LoanInfoMapper, LoanInfo> {

    public void add(LoanInfo info) {
        baseMapper.insert(info);
    }

}
