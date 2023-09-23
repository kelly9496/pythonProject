# # print(round(5.009, 2))
# # def func():
# #     a = {}
# #     print(a)
# #     return a
# #
# # print(func())
#
# list = [1, 1, 2, 3, 4, 4, 5, 5, 6, 7]
#
# set = set(list)
#
# for i in set:
#     print(i)
#
# def get_sub_set(nums):
#     # print('nums', nums)
#     # if len(nums) >=14:
#     #     nums = list(nums)[:14]
#     sub_sets = [[]]
#     for x in nums:
#         sub_sets.extend([item + [x] for item in sub_sets])
#     return sub_sets
#
# subsets = get_sub_set([1,2,3,4,5])
# subsets = [x for x in subsets if x != []]
# subsets.reverse()
# print(subsets)
#
# dic = {1:1, 2:2}
# print(len(dic))

list_a = []
list_b = [1, 3, 4]
list_a.append(list_b)
print(list_a)

if len(glData_reim_staffperM) >= 15:
    set_JE = set(glData_reim_staffperM['JE Header Id'].to_list())
    subsets_JE = get_sub_set(list(set_JE))
    subsets_glIndex_reim = list()
    for subset_JE in subsets_JE:
        glIndex_JEsubset = list(
            glData_reim_staffperM.loc[glData_reim_staffperM['JE Header Id'].isin(subset_JE)].index.values)
        subsets_glIndex_reim.append(glIndex_JEsubset)
else:
    subsets_glIndex_reim = get_sub_set(glValue_list_reim.keys())