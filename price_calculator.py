"""
Price calculation functions for CD and LP pricing based on cost.
"""

def get_cd_price(cost):
    """
    Determine the selling price for CDs based on cost.
    """
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 11.99:
        return 16.99
    elif cost <= 12.99:
        return 17.99
    elif cost <= 13.99:
        return 21.99
    elif cost <= 14.99:
        return 22.99
    elif cost <= 15.99:
        return 24.99
    elif cost <= 16.99:
        return 25.99
    elif cost <= 17.99:
        return 26.99
    elif cost <= 18.99:
        return 27.99
    elif cost <= 19.99:
        return 29.99
    elif cost <= 20.99:
        return 31.99
    else:
        return cost * 1.4  # If cost > 20.99, price = cost * 1.4

def get_lp_price(cost):
    """
    Determine the selling price for LPs based on cost.
    """
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 10.99:
        return 19.99
    elif cost <= 11.99:
        return 22.99
    elif cost <= 12.99:
        return 22.99
    elif cost <= 13.99:
        return 23.99
    elif cost <= 14.99:
        return 24.99
    elif cost <= 15.99:
        return 25.99
    elif cost <= 16.99:
        return 27.99
    elif cost <= 17.99:
        return 29.99
    elif cost <= 18.99:
        return 30.99
    elif cost <= 19.99:
        return 31.99
    elif cost <= 20.99:
        return 33.99
    elif cost <= 21.99:
        return 34.99
    elif cost <= 22.99:
        return 35.99
    elif cost <= 23.99:
        return 36.99
    elif cost <= 24.99:
        return 38.99
    elif cost <= 25.99:
        return 39.99
    elif cost <= 26.99:
        return 41.99
    elif cost <= 27.99:
        return 44.99
    elif cost <= 28.99:
        return 45.99
    elif cost <= 29.99:
        return 46.99
    elif cost <= 30.99:
        return 47.99
    elif cost <= 31.99:
        return 48.99
    elif cost <= 32.99:
        return 49.99
    elif cost <= 33.99:
        return 50.99
    elif cost <= 34.99:
        return 52.99
    elif cost <= 35.99:
        return 54.99
    elif cost <= 36.99:
        return 55.99
    elif cost <= 37.99:
        return 58.99
    elif cost <= 38.99:
        return 59.99
    elif cost <= 39.99:
        return 61.99
    elif cost <= 40.99:
        return 62.99
    elif cost <= 41.99:
        return 64.99
    elif cost <= 42.99:
        return 65.99
    elif cost <= 43.99:
        return 66.99
    elif cost <= 44.99:
        return 68.99
    elif cost <= 45.99:
        return 69.99
    elif cost <= 46.99:
        return 71.99
    elif cost <= 47.99:
        return 73.99
    elif cost <= 48.99:
        return 74.99
    elif cost <= 49.99:
        return 76.99
    else:
        return cost * 1.4  # If cost > 49.99, price = cost * 1.4