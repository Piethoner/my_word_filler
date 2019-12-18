微软win32 word接口参考: https://docs.microsoft.com/zh-CN/office/vba/api/word.document

# my_word_filler
####
使用配置的方式填充word文档
####

### 已有的配置类型
```
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
```

### 各种配置类型的参数
```
DOC_MAP = {
        # field_name : [[0, args], [1, args], [5, '', {}]]
        # [INSERT_BY_MATCH, match_text, before, underline, offset]
        # [INSERT_BY_REGEX, pattern, before, underline, offset, (flag:可有可无)]
        # [REPLACE_BY_MATCH, match_text, underline]
        # [REPLACE_BY_REGEX, pattern, underline, (flag:可有可无)]
        # [INSERT_INTO_TABLE_CELL, table_index, row, column, replace, underline]
        # [MAKE_CHOICE, match_text, options]
        # [ADD_TABLE_AFTER_CAPTION, caption, (scope:可有可无)]
        # [DEPENDENT_TYPE, 依赖的字段名称, {依赖字段值1: [INSERT_BY_MATCH, match_text, before, underline, offset],
        #                     依赖字段值2: [INSERT_BY_MATCH, match_text, before, underline, offset]}]
        # [SLICE_TYPE, (start_index, end_index), [0, match_text, before, underline, offset]] # start_index 和 end_index 为拆分部分的索引
        # [SPLIT_TYPE, (分隔符, 片段数量, index), [0, match_text, before, underline, offset]]
    }
```
