# -*- coding: utf-8 -*-
"""legacy 脚本统一路径配置。

项目根目录默认取本文件所在目录，可用环境变量 PRIVACY_DETECTIVE_BASE 覆盖。
原始输入文件（制度标尺 Excel、国标 PDF、经纪约 word 等）未随代码开源，
运行前请把对应文件放到 BASE 目录下的对应位置。
"""
import os
from pathlib import Path

BASE = os.environ.get("PRIVACY_DETECTIVE_BASE", str(Path(__file__).resolve().parent))
