def get_sub_set(nums):
    print('nums', nums)
    # if len(nums) >=14:
    #     nums = list(nums)[:14]
    sub_sets = [[]]
    for x in nums:
        sub_sets.extend([item + [x] for item in sub_sets])
    return sub_sets

list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]

print(get_sub_set(list))