# -*-coding:utf-8-*-

import copy
from word_operator import WordOperator
from doc_type import *


# 规则
def rule(wd, fcontent, args):
    if args[0] == INSERT_BY_MATCH:
        wd.insert_by_match(args[1], fcontent, args[2], args[3], args[4])
    elif args[0] == INSERT_BY_REGEX:
        flag = args[5] if len(args) > 5 else 0
        wd.insert_by_regex(args[1], fcontent, args[2], args[3], args[4], flag)
    elif args[0] == REPLACE_BY_MATCH:
        wd.replace_by_match(args[1], fcontent, args[2])
    elif args[0] == REPLACE_BY_REGEX:
        flag = args[3] if len(args) > 3 else 0
        wd.replace_by_regex(args[1], fcontent, args[2], flag)
    elif args[0] == INSERT_INTO_TABLE_CELL:
        wd.insert_into_table_cell(args[1], args[2], args[3], fcontent, args[4], args[5])
    elif args[0] == MAKE_CHOICE:
        wd.make_choice(args[1], args[2], fcontent)


# 填充文档
def fill_doc(file_path, fill_content):
    with WordOperator(file_path) as wd:
        field_map = DOC_MAP
        for field_name, args_list in field_map.items():
            fcontent = fill_content.get(field_name, '')
            fc_temp = copy.copy(fcontent)
            print(field_name, ' : ', fcontent)
            for args in args_list:
                print(args)
                if args:
                    if args[0] == DEPENDENT_TYPE:
                        args = args[2].get(fill_content.get(args[1]))
                    elif args[0] == SLICE_TYPE:
                        fcontent = copy.copy(fc_temp)
                        fcontent = fcontent[args[1][0]:args[1][1]]
                        args = args[2]
                    elif args[0] == SPLIT_TYPE:
                        fcontent = copy.copy(fc_temp)
                        fragments = fcontent.split(args[1][0])
                        if not len(fragments) == args[1][1]:
                            break
                        fcontent = fragments[args[1][2]]
                        args = args[2]

                    rule(wd, fcontent, args)




if __name__ == '__main__':
    fc = {
        u'合同标题': u'test项目test期test工程',
        u'发包方（甲方）': u'甲方test',
        u'承包方（乙方）': u'乙方test',
        u'合同订立时间': u'2019-12-12',
        u'合同订立地点': u'深圳市南山区test',
        u'合同编号': u'test123456test',
        u'合同名称': 'test项目施工合同',
        u'工程名称': u'test工程',
        u'工程地点': u'深圳市test区test路',
        u'建筑面积': u'3000平方米',
        u'开放区说明': u'开放区说明test',
        u'结构形式': u'结构形式test',
        u'图纸版号': u'test123141241test',
        u'工程内容': u'test工程内容test工程内容正文部分test工程内容正文部分2test',
        u'工程质量标准': u'验收合格率test',
        u'计税方法': u'简易计税方法',
        u'开票方纳税人资质': u'增值税一般纳税人',
        u'税率': u'6%',
        u'发票类型':u'增值税普通发票',
        u'合同计价模式': u'暂定总价',
        u'合同总价': u'10,000,000',
        u'合同总价大写': u'一千万',
        u'不含增值税合同价款': u'5,000,000',
        u'不含增值税合同价款大写': u'五百万',
        u'增值税税款': u'50,000',
        u'增值税税款大写': u'五万',
        u'水电费A/B模式': u'B模式',
        u'发票开具的两种模式选择': u'A模式',
        u'甲方代表': u'工地代表甲',
        u'乙方代表': u'工地代表乙',
        u'监理单位': u'监理单位test',
        u'总监': u'总监test',
        u'总数量': u'20',
        u'竣工图数量': u'18',
        u'总工期': u'98',
        u'开工日期': u'2020-01-20',
        u'竣工日期': u'2021-01-20',
        u'工期节点': u'',
        u'具体的交付质量标准': u'符合test标准',
        u'甲供材料/设备等范围界定': u'',
        u'材料调差': u'材料调差test',
        u'承包方式': u'模拟清单招标',
        u'计价依据': u'',
        u'预付款': u'250,000',
        u'工程款': u'',
        u'保修金': u'',
        u'保修内容': u'保修内容test',
        u'工程保修金总额': u'6',
        u'工程保修阶段甲方的代表部门': u'代表部门甲',
        u'纳税人识别码': u'test12314212412414124151test',
        u'开户行': u'中国银行深圳分行',
        u'帐号': u'test254353533tesst',
        u'地址及联系电话': u'联系地址和电话test',
        u'编制说明':u'',
        u'投标报价表': u'',
    }
    fill_doc(u'C:\\Users\\xuhuan\\Desktop\\fill_doc\\123.docx', fc)





