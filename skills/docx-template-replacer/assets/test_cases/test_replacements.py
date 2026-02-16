#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test case for docx-template-replacer skill.

This test demonstrates replacing content in a graduation project task template
while preserving all formatting.
"""

import sys
import os

# Add skill scripts to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'scripts'))
from replace_docx import replace_in_docx

# Test configuration
TEMPLATE = os.path.join(os.path.dirname(__file__), 'template.docx')
OUTPUT = os.path.join(os.path.dirname(__file__), 'test_output.docx')

# Replacement rules
REPLACEMENTS = [
    # Title
    ("基于三维激光雷达的智能移动充电桩的自动驾驶方法研究", 
     "基于ESP32的无感FOC驱动设计"),
    
    # Purpose and significance
    ("随着电动汽车和自动驾驶技术的广泛推广，充电基础设施的需求呈现指数级增长。传统的固定式充电桩在服务覆盖、充电效率以及用户交互体验方面暴露出诸多不足。智能移动充电桩作为一种突破性的技术革新，其能够自主导航至电动汽车停放点实施充电作业，极大地提高了充电服务的便捷性与效率。目前，大多数自动驾驶底盘系统依赖于二维激光雷达进行环境感知与导航，然而，该技术在应对复杂地形（如坡道、非结构化户外环境等）时表现出一定的局限性，这限制了智能移动充电桩的性能及其应用场景的拓展。本毕业设计的主要目标是针对复杂环境中自动驾驶底盘的现有缺陷，研究并开发一种基于三维激光雷达技术的智能移动充电桩自动驾驶方法，以期提升其在复杂环境下的作业能力。",
     "随着电动汽车和工业自动化的快速发展，高效、精确的电机控制技术成为关键需求。FOC（磁场定向控制）作为一种先进的电机控制算法，能够实现对永磁同步电机和交流异步电机的高性能控制，具有转矩响应快、调速范围宽、运行效率高等优点。传统的有感FOC控制需要安装位置传感器，增加了系统成本和复杂性，而无感FOC通过观测器算法估算转子位置和速度，降低了硬件成本，提高了系统可靠性。"),
    
    # Student tasks
    ("（1）系统性地调研和掌握自动驾驶底盘技术的国内外研究动态，深入理解自动驾驶的基本原理；熟练掌握Linux操作系统、机器人操作系统ROS2以及Python/C++编程语言。",
     "（1）系统性地调研无感FOC控制技术的国内外研究动态，深入理解FOC控制的基本原理；掌握ESP32开发环境、电机驱动硬件电路设计以及C/C++编程语言。"),
    
    ("（2）深入探究三维激光雷达的定位算法、地面分割算法和路径规划算法；全面了解国内外自动驾驶底盘的构型设计及其发展趋势。",
     "（2）深入研究无感FOC控制算法，包括Clarke变换、Park变换、SVPWM调制技术；掌握转子位置观测算法（如滑模观测器SMO、扩展卡尔曼滤波EKF等）。"),
    
    ("（3）完成电机驱动系统的设计，并制定自动驾驶底盘的运动控制策略，确保其运动的高效性和稳定性；实现激光雷达的驱动程序部署，确保其在rviz2可视化界面中准确输出点云图；",
     "（3）完成电机驱动硬件电路设计，包括电源模块、三相逆变桥、电流采样电路、保护电路等；实现ESP32与驱动电路的接口设计和程序开发。"),
    
    ("（4）研究并应用三维雷达激光惯性里程计定位算法fast-lio2，并在ROS2环境中进行仿真测试；开发地面分割功能包，并利用Cartographer进行地图构建；",
     "（4）基于ESP32平台实现无感FOC控制算法，包括电流环、速度环设计；完成电机参数辨识和控制器参数整定。"),
    
    ("（5）编写点云到栅格地图转换的功能包，以及全局路径规划和局部路径规划的功能包；设计并实现云服务器后端API接口；开发网页调度用户界面（UI）。",
     "（5）开发上位机监控软件，实现与ESP32的通信，能够实时显示电机运行状态（转速、电流、电压等），并支持参数在线调试。"),
    
    # Time allocation
    ("完成自动驾驶底盘选型与控制，激光雷达驱动程序开发；", 
     "完成ESP32开发环境搭建，电机驱动硬件电路设计；"),
    
    ("里程计与地图定位算法实现，地图构建功能开发；", 
     "实现SVPWM调制算法，电流环和速度环控制；"),
    
    ("路径规划算法开发与实车自动驾驶测试；", 
     "实现无感位置观测算法，电机参数辨识与调试；"),
    
    ("完成云服务器通信与网页调度界面开发，毕业论文撰写；", 
     "完成上位机软件开发，毕业论文撰写；"),
]


def main():
    """Run the test case."""
    print("=" * 60)
    print("DOCX Template Replacer - Test Case")
    print("=" * 60)
    print(f"\nTemplate: {TEMPLATE}")
    print(f"Output: {OUTPUT}")
    print(f"\nReplacements: {len(REPLACEMENTS)} pairs")
    print()
    
    if not os.path.exists(TEMPLATE):
        print(f"Error: Template file not found: {TEMPLATE}")
        sys.exit(1)
    
    # Remove old output if exists
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)
        print(f"Removed old output: {OUTPUT}")
    
    # Run replacement
    success = replace_in_docx(TEMPLATE, OUTPUT, REPLACEMENTS)
    
    if success:
        print("\n" + "=" * 60)
        print("Test completed successfully!")
        print(f"Output file: {OUTPUT}")
        print("=" * 60)
    else:
        print("\nTest failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
