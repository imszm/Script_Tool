from __future__ import annotations

import argparse
import sys
from pathlib import Path

from src import config
from src.driver import SerialDriver
from src.legacy_runner import run_legacy_script
from src.logger import setup_logging

from src.tests.test_ccb_smt import CcbSmtTest
from src.tests.test_charging import ChargingTest
from src.tests.test_horn import HornTest
from src.tests.test_pc_upgrade import PcUpgradeTest
from src.tests.test_turn_signal import TurnSignalTest
from src.tests.test_w3_power import W3PowerTest


ROOT_DIR = Path(__file__).resolve().parent
TOOL_DIR = ROOT_DIR / "Tool"


def _add_common_args(p: argparse.ArgumentParser) -> None:
    p.add_argument("-n", "--loops", type=int, default=10, help="循环次数（默认: 10）")


def _connect_driver(name: str, port: str, baudrate: int) -> SerialDriver:
    d = SerialDriver(port, baudrate, name)
    if not d.connect():
        raise SystemExit(2)
    return d


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    setup_logging().info(">>> 自动化脚本入口 main.py 启动 <<<")

    parser = argparse.ArgumentParser(
        prog="main.py",
        description="统一入口：工程化运行 Tool/ 下的脚本与 src/ 的测试用例",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    # ===== 工程化用例（src/tests）=====
    p = sub.add_parser("w3-power", help="W3 继电器开关机测试（src 版本）")
    _add_common_args(p)

    p = sub.add_parser("charging", help="继电器充电压力测试（src 版本）")
    _add_common_args(p)

    p = sub.add_parser("pc-upgrade", help="PC 升级工具压力测试（pywinauto）")
    _add_common_args(p)

    p = sub.add_parser("ccb-smt", help="CCB SMT 自动化测试（COM12 + 像素识别）")
    _add_common_args(p)

    p = sub.add_parser("turn-signal", help="左右转向灯压力测试（继电器 ASCII 指令）")
    _add_common_args(p)
    p.add_argument("--side", choices=["left", "right"], default="left", help="测试左灯或右灯")

    p = sub.add_parser("horn", help="喇叭压力测试（继电器 0x4F/0x50）")
    _add_common_args(p)

    # ===== 兼容运行（Tool/ 历史脚本）=====
    def legacy(name: str, rel_path: str, help_text: str) -> None:
        sp = sub.add_parser(name, help=help_text)
        sp.set_defaults(_legacy_rel_path=rel_path)

    legacy("relay-power-cycle", "继电器开关机压力测试.py", "继电器开关机压力测试（Tool 原脚本）")
    legacy("relay-power-cycle-enable", "继电器开关机压力测试 - 使能版本.py", "继电器开关机压力测试-使能版（Tool 原脚本）")
    legacy("relay-charge-legacy", "继电器充电压力测试.py", "继电器充电压力测试（Tool 原脚本）")
    legacy("w3-power-legacy", "W3继电器开关机压力测试.py", "W3 继电器开关机测试（Tool 原脚本）")
    legacy("relay-status", "查看继电器状态.py", "查看继电器状态（Tool 原脚本）")
    legacy("nfc-servo-stress", "总线舵机NFC压力测试.py", "总线舵机 NFC 压力测试（Tool 原脚本）")
    legacy("nfc-servo-ini-stress", "舵机demo.py", "NFC 舵机 INI 驱动压力测试（Tool 原脚本）")
    legacy("nfc-card-stat", "NFC刷卡统计调试版本1.0.py", "NFC 刷卡统计调试（Tool 原脚本）")
    legacy("ota-gui", "OTA升级工具优化3（优化升级失败继续升级）.py", "OTA 循环升级自动化 GUI（Tool 原脚本）")
    legacy("upgrade-stress-screenshot", "升级工具压力自动化测试 - 可截图版本.py", "升级工具压力测试-截图版（Tool 原脚本）")
    legacy("assembly-tool", "组装生产工具压力测试.py", "组装生产工具压力测试（Tool 原脚本）")
    legacy("after-sales-tool", "售后工具脚本测试.py", "售后工具脚本测试（Tool 原脚本）")
    legacy("turn-signals-legacy", "左右转向灯自动化测试-正式版本.py", "左右转向灯测试（Tool 原脚本）")
    legacy("mouse-pos", "鼠标定位脚本.py", "鼠标定位脚本（Tool 原脚本）")
    legacy("time-diff", "build/时间差计算.py", "时间差计算工具（Tool 原脚本）")
    legacy("testcases-excel", "测试用例.py", "生成 Excel 测试用例（Tool 原脚本）")
    legacy("ble-lrd", "LRD调试程序.py", "BLE DLL 串口调试（Tool 原脚本）")

    args = parser.parse_args(argv)

    # legacy scripts: run and exit
    if hasattr(args, "_legacy_rel_path"):
        run_legacy_script(TOOL_DIR / getattr(args, "_legacy_rel_path"))
        return 0

    drivers: dict[str, SerialDriver] = {}
    try:
        if args.cmd == "w3-power":
            drivers["relay"] = _connect_driver("RelayGen", config.RELAY_PORT, config.BAUDRATE_RELAY)
            drivers["device"] = _connect_driver("Device", config.DEVICE_PORT, config.BAUDRATE_DEVICE)
            W3PowerTest(drivers).run(args.loops)
            return 0

        if args.cmd == "charging":
            drivers["relay"] = _connect_driver("RelayGen", config.RELAY_PORT, config.BAUDRATE_RELAY)
            drivers["device"] = _connect_driver("Device", config.DEVICE_PORT, config.BAUDRATE_DEVICE)
            ChargingTest(drivers).run(args.loops)
            return 0

        if args.cmd == "pc-upgrade":
            PcUpgradeTest(drivers).run(args.loops)
            return 0

        if args.cmd == "ccb-smt":
            drivers["relay"] = _connect_driver("RelayCCB", config.RELAY_CCB_PORT, config.BAUDRATE_RELAY)
            CcbSmtTest(drivers).run(args.loops)
            return 0

        if args.cmd == "turn-signal":
            drivers["relay"] = _connect_driver("RelayGen", config.RELAY_PORT, config.BAUDRATE_RELAY)
            TurnSignalTest(drivers).run(args.loops, side=args.side)
            return 0

        if args.cmd == "horn":
            drivers["relay"] = _connect_driver("RelayGen", config.RELAY_PORT, config.BAUDRATE_RELAY)
            HornTest(drivers).run(args.loops)
            return 0

        raise SystemExit(2)
    finally:
        for d in drivers.values():
            d.close()


if __name__ == "__main__":
    raise SystemExit(main())