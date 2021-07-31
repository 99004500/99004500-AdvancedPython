from main import rownumber, sheet1


def test_rownumber1():
    assert rownumber(sheet1, 88002500) == 2
    assert rownumber(sheet1, 88002501) == 3

def test_rownumber2():
    assert rownumber(sheet1, 88002502) == 4
    assert rownumber(sheet1, 88002503) == 5

def test_rownumber3():
    assert rownumber(sheet1, 88002504) == 6
    assert rownumber(sheet1, 88002505) == 7

def test_rownumber4():
    assert rownumber(sheet1, 88002506) == 8
    assert rownumber(sheet1, 88002507) == 9

def test_rownumber5():
    assert rownumber(sheet1, 88002508) == 10
    assert rownumber(sheet1, 88002509) == 11
    assert rownumber(sheet1, 88002510) == 12

