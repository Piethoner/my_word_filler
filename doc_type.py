# -*-coding:utf-8-*-

import re

INSERT_BY_MATCH = 0
INSERT_BY_REGEX = 1
REPLACE_BY_MATCH = 2
REPLACE_BY_REGEX = 3
INSERT_INTO_TABLE_CELL = 4
MAKE_CHOICE = 5
ADD_TABLE_AFTER_CAPTION = 6
# 有些字段的填写位置依赖其他字段的值
DEPENDENT_TYPE = 15
# 有些字段需要拆分为多个部分填写到多个位置
SLICE_TYPE = 16
SPLIT_TYPE = 17





DOC_MAP = {
        # field_name : [[0, args], [1, args], [5, '', {}]]
        # [INSERT_BY_MATCH, match_text, before, underline, offset]
        # [INSERT_BY_REGEX, pattern, before, underline, offset, (flag:可有可无)]
        # [REPLACE_BY_MATCH, match_text, underline]
        # [REPLACE_BY_REGEX, pattern, underline, (flag:可有可无)]
        # [INSERT_INTO_TABLE_CELL, table_index, row, column, replace, underline]
        # [MAKE_CHOICE, match_text, options]
        # [ADD_TABLE_AFTER_CAPTION, caption, (scope:可有可无)]
        # [DEPENDENT_TYPE, 依赖的字段名称, {依赖字段值1: [0, match_text, before, underline, offset],
        #                     依赖字段值2: [0, match_text, before, underline, offset]}]
        # [SLICE_TYPE, (start_index, end_index), [0, match_text, before, underline, offset]] # start_index 和 end_index 为拆分部分的索引
        # [SPLIT_TYPE, (分隔符, 片段数量, index), [0, match_text, before, underline, offset]]
        u'合同标题': [
                [REPLACE_BY_MATCH, u'××项目×期×工程', False],
                [REPLACE_BY_MATCH, u'××万科××项目×期×标段工程', False],
                [REPLACE_BY_MATCH, u'杭州万科**项目装饰装修工程施工合同', False]
        ],
        u'发包方（甲方）': [
                [INSERT_INTO_TABLE_CELL, 1, 1, 2, True, True],
                [REPLACE_BY_REGEX, r'发包方(?:\(|（)以下简称甲方(?:\)|）)：(.*)', True],
                [REPLACE_BY_REGEX, r'甲方（全称）：(.*)', True],
                [INSERT_BY_MATCH, u'甲方（采购方全称）：', False, True, 0],
                [REPLACE_BY_REGEX, r'甲方：( +)(?!.{,100}?(?:\(|（)公章(?:\)|）))', True, re.S],
                [REPLACE_BY_REGEX, r'承 {,2}诺 {,2}函\n(.*?\n?.*?)\n对于贵公司与我公司', True],
                # [SPLIT_TYPE, (u'市', 2, 0), [REPLACE_BY_REGEX,
                #                             r'发包方(?:\(|（)以下简称甲方(?:\)|）)：( +?)市万科房地产有限公司', True]]
        ],
        u'承包方（乙方）': [
                [INSERT_INTO_TABLE_CELL, 1, 2, 2, True, True],
                [INSERT_BY_MATCH, u'乙方（供应方全称）：', False, True, 0],
                [REPLACE_BY_REGEX, r'承包方(?:\(|（)以下简称乙方(?:\)|）)：( +?)(?:公司)?', True],
                [REPLACE_BY_REGEX, r'乙方（全称）：(.*)', True],
                [REPLACE_BY_REGEX, r'乙方：( +)(?!.{,100}?(?:\(|（)公章(?:\)|）))', True, re.S],
                [INSERT_BY_MATCH, u'单位名称：', False, False, 0]
        ],
        u'合同订立时间': [
                [INSERT_INTO_TABLE_CELL, 1, 3, 2, True, True],
                [SLICE_TYPE, (0, 4), [REPLACE_BY_REGEX, r'甲乙双方于(.*?)年', True]],
                [SLICE_TYPE, (5, 7), [REPLACE_BY_REGEX, r'甲乙双方于(?:.*?)年(.*?)月', True]],
                [SLICE_TYPE, (8, 10), [REPLACE_BY_REGEX, r'甲乙双方于(?:.*?)年(?:.*?)月(.*?)日', True]],
                [SLICE_TYPE, (0, 4), [REPLACE_BY_REGEX, r'双方经协商于(.*?)年', True]],
                [SLICE_TYPE, (5, 7), [REPLACE_BY_REGEX, r'双方经协商于(?:.*?)年(.*?)月', True]],
                [SLICE_TYPE, (8, 10), [REPLACE_BY_REGEX, r'双方经协商于(?:.*?)年(?:.*?)月(.*?)日', True]]
        ],
        u'合同订立地点': [
                [INSERT_INTO_TABLE_CELL, 1, 4, 2, True, True]
        ],
        u'合同编号': [
                [INSERT_INTO_TABLE_CELL, 1, 5, 2, True, True]
        ],
        u'合同名称': [
                [REPLACE_BY_REGEX, r'《(.*?项目总包工程施工合同)》', True],
                [INSERT_BY_MATCH, u'合同（以下简称原合同）', True, True, 0],
                [REPLACE_BY_REGEX, r'对于贵公司与我公司就(?:.*?)项目签订的《(.*?)(?:施工总承包合同)?》', True],
                [REPLACE_BY_REGEX, r'双方经协商于(?:.*?)年(?:.*?)签订了(.*?)合同', True]
        ],
        u'工程名称': [
                [REPLACE_BY_REGEX, r'工程名称：(.+)', True],
                [REPLACE_BY_REGEX, r'对于贵公司与我公司就(.*?)项目签订', True],
                [REPLACE_BY_REGEX, r'我公司已与在贵公司(.*?)项目施工', True],
                [REPLACE_BY_REGEX, r'若在(.+?)项目', True],
                [REPLACE_BY_REGEX, r'对( +?)(?:（工程全称）)', True],
        ],
        u'工程地点': [
                [REPLACE_BY_REGEX, u'工程地点：(.+)', True],
        ],
        u'建筑面积': [
                [INSERT_BY_MATCH, u'1.3 建筑面积：约', False, True, 0]
        ],
        u'开放区说明': [
                [REPLACE_BY_REGEX, r'开放区说明:(.*)', True]
        ],
        u'结构形式': [
                [INSERT_BY_MATCH, u'结构形式：', False, True, 0]
        ],
        u'图纸版号': [
                [INSERT_BY_MATCH, u'甲方将委托乙方承担图纸版号为', False, True, 0]
        ],
        u'工程内容': [
                [REPLACE_BY_REGEX, u'所包含之(.*?)。详细见附件《总包及分包界面划分表》。', True]
        ],
        u'工程质量标准': [
                [INSERT_BY_MATCH, u'本工程质量标准为：', False, True, 0]
        ],
        u'开票方纳税人资质': [
                [MAKE_CHOICE, u'增值纳税人类型及计税方法', (u'增值税一般纳税人', u'增值税小规模纳税人')]
        ],
        u'计税方法':[
                [DEPENDENT_TYPE, u'开票方纳税人资质', {
                        u'增值税一般纳税人':
                                [MAKE_CHOICE, u'[□] ?增值税一般纳税人', (u'一般计税方法', u'简易计税方法')],
                        u'增值税小规模纳税人':
                                [MAKE_CHOICE, u'[□] ?增值税小规模纳税人', (u'简易计税方法',)]
                }]
        ],
        u'发票类型': [
                [MAKE_CHOICE, u'开具发票类型及适用税率或征收率', \
                        (u'增值税专用发票', u'增值税普通发票', u'除增值税专用发票以外的其他增税扣税凭证')]
        ],
        u'税率': [
                [DEPENDENT_TYPE, u'发票类型', {
                        u'增值税专用发票':
                                [MAKE_CHOICE, u'[□] ?增值税专用发票', (u'13%', u'9%', u'6%', u'3%')],
                        u'增值税普通发票':
                                [MAKE_CHOICE, u'[□] ?增值税普通发票', (u'13%', u'9%', u'6%', u'3%')],
                        u'除增值税专用发票以外的其他增税扣税凭证':
                                [MAKE_CHOICE, u'[□] ?除增值税专用发票以外的其他增税扣税凭证', (u'13%', u'9%', u'6%', u'3%', u'0%')],
                }]
        ],
        u'合同计价模式': [
                [MAKE_CHOICE, u'合同计价模式为：', (u'固定总价', u'暂定总价')]
        ],
        u'合同总价': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：￥(.*?)元', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：￥(.*?)元', True]
                }]
        ],
        u'合同总价大写': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：￥(?:.*?)元，大写：人民币(.*?)圆整', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：￥(?:.*?)元，大写：人民币(.*?)圆整', True]
                }]
        ],
        u'不含增值税合同价款': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：(?:.*?)其中不含增值税合同价款：￥(.*?)元', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：(?:.*?)其中不含增值税合同价款：￥(.*?)元', True]
                }]
        ],
        u'不含增值税合同价款大写': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：(?:.*?)其中不含增值税合同价款：￥(?:.*?)元，大写：人民币(.*?)圆整', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：(?:.*?)其中不含增值税合同价款：￥(?:.*?)元，大写：人民币(.*?)圆整', True]
                }]
        ],
        u'增值税税款': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：(?:.*?)增值税税款￥(.*?)元', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：(?:.*?)增值税税款￥(.*?)元', True]
                }]
        ],
        u'增值税税款大写': [
                [DEPENDENT_TYPE, u'合同计价模式', {
                        u'固定总价': [REPLACE_BY_REGEX, r'合同固定总价：(?:.*?)增值税税款￥(?:.*?)元，大写：人民币(.*?)圆整。', True],
                        u'暂定总价': [REPLACE_BY_REGEX, r'合同暂定总价：(?:.*?)增值税税款￥(?:.*?)元，大写：人民币(.*?)圆整。', True]
                }]
        ],
        u'水电费A/B模式': [
                [MAKE_CHOICE, u'现场水电费计量采用', (u'A模式', u'B模式')]
        ],
        u'发票开具的两种模式选择': [
                [MAKE_CHOICE, u'发票开具的两种模式选择：', (u'A模式', u'B模式')]
        ],
        u'甲方代表': [
                [REPLACE_BY_REGEX, r'甲方驻工地总?代表为(.*?)。', True]
        ],
        u'乙方代表': [
                [REPLACE_BY_REGEX, r'乙方驻工地(?:(?:代表)|(?:现场经理))为(.*?)。', True]
        ],
        u'监理单位': [
                [REPLACE_BY_REGEX, r'监理单位：(.*?)，', True]
        ],
        u'总监': [
                [REPLACE_BY_REGEX, r'总监：(.*?)指监理单', True]
        ],
        u'总数量': [
                [REPLACE_BY_REGEX, r'甲方向乙方提供图纸(.*?)套', True]
        ],
        u'竣工图数量': [
                [REPLACE_BY_REGEX, r'制作竣工图所用图纸(.*?)套', True]
        ],
        u'总工期': [
                [REPLACE_BY_REGEX, r'工期：约(.*?)天', True],
                [REPLACE_BY_REGEX, r'本合同工程总工期为(.*?)日历天', True],
                [REPLACE_BY_REGEX, r'全部竣工共(.*?)天', True]
        ],
        u'开工日期': [
                [SLICE_TYPE, (0, 4), [REPLACE_BY_REGEX, r'开工日期暂定为?(.*?)年', True]],
                [SLICE_TYPE, (5, 7), [REPLACE_BY_REGEX, r'开工日期暂定为?(?:.*?)年(.*?)月', True]],
                [SLICE_TYPE, (8, 10), [REPLACE_BY_REGEX, r'开工日期暂定为?(?:.*?)年(?:.*?)月(.*?)日', True]],
                [REPLACE_BY_REGEX, r'工期安排：暂定为?( +)', True]

        ],
        u'竣工日期': [
                [SLICE_TYPE, (0, 4), [REPLACE_BY_REGEX, r'竣工日期(.*?)年', True]],
                [SLICE_TYPE, (5, 7), [REPLACE_BY_REGEX, r'竣工日期(?:.*?)年(.*?)月', True]],
                [SLICE_TYPE, (8, 10), [REPLACE_BY_REGEX, r'竣工日期(?:.*?)年(?:.*?)月(.*?)日', True]],
                [SLICE_TYPE, (0, 4), [REPLACE_BY_REGEX, r'进场，要求(.*?)年', True]],
                [SLICE_TYPE, (5, 7), [REPLACE_BY_REGEX, r'进场，要求(?:.*?)年(.*?)月', True]],
                [SLICE_TYPE, (8, 10), [REPLACE_BY_REGEX, r'进场，要求(?:.*?)年(?:.*?)月(.*?)日', True]]
        ],
        u'工期节点': [
                []
        ],
        u'具体的交付质量标准': [
                [INSERT_BY_MATCH, r'具体的交付质量标准：施工质量', False, True, 0]
        ],
        u'甲供材料/设备等范围界定': [
                []
        ],
        u'材料调差': [
                [REPLACE_BY_REGEX, r'材料调差：(.*)', False]
        ],
        u'承包方式': [
                [MAKE_CHOICE, u'本招标工程的承包方式：', (u'费率招标', u'模拟清单招标', u'工程量清单招标')]
        ],
        u'计价依据': [
                []
        ],
        u'预付款': [
                [INSERT_BY_MATCH, u'预付款为', False, True, 0]
        ],
        u'工程款': [
                [REPLACE_BY_REGEX, r'工程款拨付累计达合同总造价的(.*?)时', False]
        ],
        u'保修金': [
                [REPLACE_BY_REGEX, r'本工程预留结算总价的(.*?)为工程质量保修款', False],
                [REPLACE_BY_REGEX, r'工程保修金为本工程结算总价（不含支付方式差异贴息部分）的(.*?)，保修金无利息', True]
        ],
        u'保修内容': [
                [REPLACE_BY_REGEX, r'保修内容：(.*)', True]
        ],
        u'工程保修金总额': [
                [REPLACE_BY_REGEX, r'不含支付方式差异补贴）的(.*?)%计取', True],
                [REPLACE_BY_REGEX, r'工程保修金总额：合同结算总价(.*?)%', True],
        ],
        u'工程保修阶段甲方的代表部门': [
                [REPLACE_BY_REGEX, r'本工程保修阶段甲方的代表部门为(.*?)，乙', True],
                [REPLACE_BY_REGEX, r'乙方应接受( +?)的管理', True],
        ],
        u'纳税人识别码': [
                [INSERT_BY_MATCH, u'纳税人识别码（社会信用证统一代码）：', False, True, 0]
        ],
        u'开户行': [
                # [SPLIT_TYPE, (u'银行', 2, 0), [REPLACE_BY_REGEX, r'开户行：(.*?)银行(?:.*?)支行', True]],
                # [SPLIT_TYPE, (u'银行', 2, 1), [REPLACE_BY_REGEX, r'开户行：(?:.*?)银行(.*?)支行', True]],
                [REPLACE_BY_REGEX, r'开户行：(.*)', False]
        ],
        u'帐号': [
                [REPLACE_BY_REGEX, r'开户行(?:.*?)\n(?:.*?)帐号：(.*)', False]
        ],
        u'地址及联系电话': [
                [INSERT_BY_REGEX, r'地址及联系电话(?::|：)?', False, False, 0]
        ],
        u'编制说明': [
                [ADD_TABLE_AFTER_CAPTION, u'经济标编制要求']
        ],
        u'投标报价表': [
                [ADD_TABLE_AFTER_CAPTION, u'工程量清单投标报价表']
        ],
}