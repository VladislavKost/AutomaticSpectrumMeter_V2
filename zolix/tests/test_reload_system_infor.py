from zolix.tests import test_common


def test():
    m = test_common.zolix_gateway.reload_system_infor()
    test_common.print_result("ReloadSystemInfor", m)


if __name__ == "__main__":
    test()
