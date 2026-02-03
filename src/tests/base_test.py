# src/tests/base_test.py
import logging

class BaseTest:
    def __init__(self, drivers):
        """
        :param drivers: 这是一个字典，包含 'relay': driver_obj, 'device': driver_obj
        """
        self.drivers = drivers
        self.logger = logging.getLogger(f"AutoTest.Case.{self.__class__.__name__}")
        self.stop_flag = False

    def setup(self):
        """测试前准备 (如: 复位继电器)"""
        pass

    def run(self, loops):
        """核心逻辑，子类必须覆盖它"""
        raise NotImplementedError

    def teardown(self):
        """测试后清理"""
        pass