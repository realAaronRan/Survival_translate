#! /usr/bin/python
# -*- coding:UTF-8 -*-
'''
@Description: This is a script for translating text.
@Author: Aaron Ran
@Github: https://github.com/realAaronRan
@FilePath: \Survival_translate\translate.py
@Date: 2020-03-03 16:25:12
@LastEditTime: 2020-03-08 21:25:53
'''

import re

def keep_the_same_as(weapon_name):
    '''
    @description: 判断武器名是否需要改变。
                对于由字母、数字和特殊字符'-'组成的武器名，不需要翻译。
    @param {str} weapon_name: 武器名字符串。 
    @return: {bool} 满足条件返回true，否则返回false。
    '''    

    for ch in weapon_name:
        if ord(ch) < ord(' ') or ord(ch) > ord('~'):
            return False
    return True

def translate(sht_raw, sht_trans, roles, return_dict):
    '''
    description: 完成翻译工作的脚本。

    Params:
        raw_file{str}: 原始文档的路径。
        trans_file{str}: 翻译文档的路径。 
        column_base_raw{int}: 基本语言在原始文件中的列索引，1，2...。 
        column_foraign_raw{int}: 需要翻译的外文在原始文档中的列索引，1,2...。 
        column_foraign_trans{int}: 需要翻译的外文在翻译文档中的列索引，1,2...。
        roles{dict}: 游戏角色性别列表。
    return: 

    '''
    # 建立映射规则表
    current_row = 0
    trans_dict = {}
    chinese_names = roles['male'] + (roles['female'])
    names_map = {}
    for current_row in range(sht_trans.shape[0]):
        # cell_value = sht_dict.range(current_row, 1).value
        cell_value = sht_trans.iloc[current_row, 0]
        # cell_value = sht_dict.range('A'+str(current_row)).value

        key = cell_value
        trans_value = sht_trans.iloc[current_row, 1]

        # 判断是否存在数字或名字
        contain_digit = re.search(r'\d+', cell_value)
        contain_name = None
        for name in chinese_names:
            if re.search(name, cell_value):
                contain_name = name
                if name == cell_value:
                    names_map[cell_value] = trans_value
                break
        
        # 只包含数字，不包含名字的情况
        if contain_digit and not contain_name:
            digit = contain_digit.group(0)
            key = cell_value.replace(digit, '$')
            quality = 'single' if digit == '1' else 'multiple'
            trans_dict.setdefault(key, {})[quality] = trans_value

        # 只包含名字，不包含数字的情况
        if contain_name and not contain_digit:
            key = cell_value.replace(contain_name, '&')
            _gender = 'male' if contain_name in roles['male'] else 'female'
            trans_dict.setdefault(key, {})[_gender] = trans_value
        
        # 同时包含名字和数字的情况
        if contain_digit and contain_name:
            digit = contain_digit.group(0)
            quality = 'single' if digit == '1' else 'multiple'
            key = cell_value.replace(digit, '$').replace(contain_name, '&')
            _gender = 'male' if contain_name in roles['male'] else 'female'
            trans_dict.setdefault(key, {})[quality+_gender] = trans_value
        
        if not contain_digit and not contain_name:
            trans_dict[key] = trans_value
    foraign_names = [f_name for f_name in names_map.values()]

    result = []
    # 主循环，遍历原始文本表的“chiese"列
    for current_row in range(1, sht_raw.shape[0]):
        cell_value = sht_raw.iloc[current_row, 1]
        if cell_value is None:
            break
        # 对武器特别处理
        type_ = sht_raw.iloc[current_row, 0]
        # 如果武器名只包含英文和数字，不需要翻译
        if type_ == '武器名' and  keep_the_same_as(cell_value):
            result.append(cell_value)
            continue

        # 判断是否存在数字或名字
        contain_digit = re.search(r'\d+', cell_value)
        contain_name = None
        for name in chinese_names:
            if re.search(name, cell_value):
                contain_name = name
                break
        
        try:
            # 只包含数字，不包含名字
            if contain_digit and not contain_name:
                digit = contain_digit.group(0)
                quality = 'single' if digit == '1' else 'multiple'
                key = cell_value.replace(digit, '$')
                trans_value = re.sub(r'\d+', digit, trans_dict[key][quality])
                result.append(trans_value)
                continue
                # trans_cell.value = trans_value

            # 只包含名字，不包含数字
            if contain_name and not contain_digit:
                _gender = 'male' if contain_name in roles['male'] else 'female'
                key = cell_value.replace(contain_name, '&')
                trans_value = trans_dict[key][_gender]
                f_name = names_map[contain_name]
                pattern_name = None
                for name in foraign_names:
                    if re.search(name, trans_value):
                        pattern_name = name
                        break
                result.append(trans_value.replace(pattern_name, f_name))
                continue
                # trans_cell.value = trans_value.replace(pattern_name, f_name)

            # 同时包含名字和数字
            if contain_name and contain_digit:
                digit = contain_digit.group(0)
                quality = 'single' if digit == '1' else 'multiple'
                _gender = 'male' if contain_name in roles['male'] else 'female'
                key = cell_value.replace(contain_name, '&').replace(digit, '$')
                trans_value = trans_dict[key][quality+_gender]
                f_name= names_map[contain_name]
                pattern_name = None
                for name in foraign_names:
                    if re.search(name, trans_value):
                        pattern_name = name
                        break
                trans_value = re.sub(r'\d+', digit, trans_value)
                trans_value = trans_value.replace(pattern_name, f_name)
                result.append(trans_value)
                continue
                # trans_cell.value = trans_value
            
            if not contain_digit and not contain_name:
                key = cell_value
                trans_value = trans_dict[key]
                result.append(trans_value)
                continue
                # trans_cell.value = trans_value
        except KeyError:
            result.append('$ERROR$')
            continue
            # trans_cell.color = (200, 70, 50)
    foraign_lang = sht_trans.columns.values[1]
    return_dict[foraign_lang] = result
    return return_dict

# if __name__ == "__main__":
#     raw_file = pd.read_excel('raw_text.xls', sheet_name='Sheet1')
#     trans_file = pd.read_excel('translate_Por_Fre_Ger.xlsx', sheet_name='Sheet1')
#     sht_raw = raw_file.loc[:, ['类型','简体中文']]
#     sht_trans = trans_file.loc[:, ['简体中文', '德语']]

#     roles = {'male': ['杰瑞', '安德森', '威廉姆斯', '泰勒', '亚伯拉罕', '达里尔', '米琼恩', '戴尔'],
#             'female': ['玛丽安娜', '夏娜', '蔷薇', '希维尔']}
#     return_dict = {}
#     result = translate(sht_raw, sht_trans, roles, return_dict)
