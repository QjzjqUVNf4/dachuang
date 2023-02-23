from typing import List, Tuple, Any

categories: list[tuple[str, list[str | Any]] | tuple[str, list[str]]] = [
        ('HC',
         ['薪', '差旅', '差费费', '办公', '保险', '公积金', '劳务', '福利', '股权', '股票', '股份', '期权', '社保',
          '人工', '人力', '人员', '工资', '团队建设',
          '学会协会会费', '员工', '职工', '人力资源', '餐费', '工会经费', '住房费用', '人才', '人事', '保障金', '奖金',
          '五险一金', '劳保', '劳动保护', '社会保障费用',
          '招聘费', '培训费', '伙食费', '工作经费', '用工费', '教育培训', '社会统筹', '课程制作费']),
        ('Mat',
         ['料', '直接投入', '物流', '制造', '耗材', '直接费用', '物资', '存货', '物耗', '易耗', '零配件', '配件',
          '消耗品', '半成品', '库存商品', '包装物', '领用',
          '资源支出', '配件费', '采购支出', '研发用消耗', '加工', '能耗', '施工费', '工程费', '燃动', '直接消耗费',
          '天然气', '电', '水费', '蒸汽费', '掩膜费', '光罩',
          '动燃', '工具使用费', '印刷费', '工具费', '工器具', '制版费', '动力', '能源', '汽油费', '水费', '动能费',
          '工装', '机械费', '开发消耗费', '木模费', '采集费用', '制造费用']),
        ('Ext',
         ['委外', '外委', '委托外部', '外包', '外协', '鉴定', '咨询', '合作', '信息', '委托', '中介', '咨询', '托管',
          '知识产权', '专业服务', '会务费',
          '第三方', '联合开发', '外部', '审']),
        ('FAD',
         ['折旧', '摊', '维', '租', '设备', '暖', '安装', '模具', '固定资产', '设施', '修理', '科研仪器', '工艺装备',
          '物业', '装修', '仪器', '装备', '房屋',
          '机房', '房费', '计量费用', '资产', '修缮', '场地', '场所', '改造']),
        ('Exp',
         ['样品', '送样', '打样', '样版', '试片', '样板', '样件模', '样机', '样车费', '调样费', '样本', '样件', '检测',
          '检验', '试验', '测试', '试产', '中试', '实验',
          '试制', '中期试制', '中间试用', '参试', '检定费', '检化验费', '专利', '成果', '出版', '版权', '技术', '商标',
          '软著', '联合开发', '专用费', '测验',
          '工艺', '管理', '评估', '验收', '服务', '注册', '审批', '调试', '备件', '行政', '制作', '设计', '认证',
          '定标费', '特许', '许可', '报批', '证书', '评审']),
        ('Oth',
        ['其他', '研究与应用', '公司', '招待', '一种', '项目', '基于', '研发费用', '会议', '交通', '邮寄', '通讯',
         #'其它', '流量费', '网络费', '带宽耗用', '公杂费','展会费',
          #'相关费用', '其余各明细', '环保', '车辆', '快递费', '保洁排污费', '运杂费', '宽带', '邮寄', '汽车', '整车',
          #'研发电话费', '保安保洁费', '治安保卫费', '各种规费',
          #'稿费', '运输', '菌种特许权使用费', '安全', '部门', '税金', '日常', '云服务器费用', '会员费', '法务费',
          #'手机费', '财务', '课题', '模特试装费', '汽费', '协会费',
          #'数据使用费', '接待费', '公告费', '新品种培育费', '运费', '事务', '精品补贴特设书店', '交际', '招持',
          #'晒图费', '车船使用费', '光罩', '安保', '税费', '开发', '调研', '小车', '临床和注册费', '花稿样板费', '引智资助经费',
          #'外采软硬件及服务', '研发临床费', '录制', '外事费', '装卸费', '手续费', '杂项', '利息费用', '光纤费',
          #'辅助费', '研发产品销售收入', '用车费', '激励太景医药研发费', '警卫消防费', '零星费用', '仓储', '临床',
          #'美术', '学会协会费用', '机械费', '一致', '通信', '临床','数据库使用费', '网站', '广告',
          #'邮递费', '示范及品种选育费', '消耗费用', '减：', '冲回', 'IT', '软件', '应酬', '系统', '平台', '市场', '运营', '专项',
          #'专业', '业务', '中心', '费用化', '资本化', '间接', '直接'
         ]),
        # ('Unclassified', ['宣传', '专家', '顾问', '指导', '产品',]),
        ('All', ['合计'])
    ]