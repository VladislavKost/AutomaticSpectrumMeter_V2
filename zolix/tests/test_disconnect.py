from zolix.tests import test_common


def test():
    m = test_common.zolix_gateway.disconnect()
    test_common.print_result("DisConnect", m)


if __name__ == "__main__":
    test()
