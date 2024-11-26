from decimal import Decimal, ROUND_HALF_UP

def my_rounding(x):
    y = Decimal(str(x)).quantize(Decimal("0.01"), rounding = ROUND_HALF_UP)
    # print(x, y)
    return y

a = Decimal('3095.575').quantize(Decimal("0.01"), rounding = ROUND_HALF_UP)
# a = round(3095.575,2)
b = my_rounding(3095.575)
print(b)


