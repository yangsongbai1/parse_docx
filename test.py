# 创建一个字典
from collections import OrderedDict

my_dict = {'apple': 10, 'banana': 5, 'orange': 8}


def update_dict(my_dict, key_to_insert_after, new_dict):
    keys = list(my_dict.keys())
    my_dict_list = list(my_dict.items())
    index = keys.index(key_to_insert_after)
    result_dict = dict(my_dict_list[:index + 1] + list(new_dict.items()) + my_dict_list[index + 1:])
    return result_dict


# 打印原始字典
print("原始字典:", my_dict)

# 指定元素后边添加新元素
key_to_insert_after = 'banana'
new_key = 'grape'
new_value = 12


dd = update_dict(my_dict, key_to_insert_after, {new_key: new_value})
print(dd)

exit()


# 判断指定的键是否在字典中
if key_to_insert_after in my_dict:
    # 使用字典的 update 方法添加新元素
    new_items = {new_key: new_value}
    temp_dict = {key: my_dict[key] for key in my_dict.keys()}
    temp_dict.update(new_items)

    # 获取指定键的索引位置
    index = list(temp_dict.keys()).index(key_to_insert_after)

    # 将新元素插入到指定位置后
    result_dict = OrderedDict(
        list(temp_dict.items())[:index + 1] + list(new_items.items()) + list(temp_dict.items())[index + 1:])

    # 打印修改后的字典
    print("修改后的字典:", result_dict)
else:
    print(f"键 '{key_to_insert_after}' 不存在于字典中。")
