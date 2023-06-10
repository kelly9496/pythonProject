from itertools import combinations
from functools import reduce

a = [[1], [2], [3], [4, 5]]
# for i in range(len(a)+1):
#     comb = combinations(a, i)
#     #print(list(comb))
# print(list(comb))
def appendTest(list, item):
    list.append(item)
    return list
newlist = reduce(lambda x,y: appendTest(x, y[0]), a, [])
print(newlist)
