## 题目

基于ESP32的FOC控制器设计

## 课题的背景和意义

随着新能源汽车、工业机器人等领域的快速发展，高效、可靠的电机控制技术成为重要研究课题。无传感器FOC（Field-Oriented Control，磁场定向控制）技术因其高效率、低噪音和宽调速范围等优点，在无刷直流电机驱动中应用广泛。

ESP32微控制器因其集成度高、功能丰富、成本低廉等特点，在工业控制、物联网、新能源汽车等领域得到广泛应用。本设计以ESP32为控制核心，研究并实现FOC算法在三相无刷直流电机控制中的应用。

本课题的主要意义在于：
1. 掌握FOC矢量控制的理论基础和算法实现
2. 深入理解ESP32微控制器的硬件特性和固件开发方法
3. 设计和实现工业级FOC控制器的硬件电路和控制算法
4. 为新能源和工业自动化领域的电机控制系统提供参考

## 毕业设计的主要内容

(1) FOC控制理论研究：系统学习三相异步电机和无刷直流电机的数学模型，深入理解FOC算法的基本原理和实现方法，包括Clark变换、Park变换、SVPWM调制等关键技术。

(2) 电机参数辨识与仿真：进行电机参数辨识实验，获取电机的抗阻、漏感、励磁感应等参数，建立电机的仿真模型，在Matlab/Simulink中验证FOC控制器的性能。

(3) 硬件电路设计：设计基于ESP32的FOC控制器硬件电路，包括电源模块、驱动模块、采样模块、通信接口等部分，确保系统的稳定性和可靠性。

(4) 控制算法实现：在ESP32上编程实现FOC控制算法，包括速度环、电流环、磁链观测器等控制策略，实现无传感器控制。

(5) 系统集成与测试：完成电机驱动系统的集成，进行空载运行、负载测试、效率测试等实验，验证控制系统的性能指标。

(6) 论文撰写与答辩：撰写毕业设计论文，制作答辩演示文稿，完成答辩演讲。

## 主要参考文献

[1] Chen We, Hongfang Lv. Design of Integrated Stepper Motor Field Oriented Control and Drive Based on ESP32-S3[C]. 2025.

[2] Zheng Wang, Jian Chen, Ming Cheng, et al. Field-Oriented Control and Direct Torque Control for Paralleled VSIs Fed PMSM Drives With Variable Switching Frequencies[J]. IEEE Transactions on Power Electronics, 2015, 31(3): 2417-2428.

[3] Bin Wu, Mehdi Narimani. High-Power Converters and AC Drives[C]. 2016.

[4] Shaohua Chen, Gang Liu, Lianqing Zhu. Sensorless Control Strategy of a 315 kW High-Speed BLDC Motor Based on a Speed-Independent Flux Linkage Function[J]. IEEE Transactions on Industrial Electronics, 2017, 64(11): 8607-8617.

[5] Joon Sung Park, Ki-Doek Lee, Sung Gu Lee, et al. Unbalanced ZCP Compensation Method for Position Sensorless BLDC Motor[J]. IEEE Transactions on Power Electronics, 2018, 34(4): 3020-3024.

[6] Cong Gu, Xiaolin Wang, Xiaoqing Shi, et al. A PLL-Based Novel Commutation Correction Strategy for a High-Speed Brushless DC Motor Sensorless Drive System[J]. IEEE Transactions on Industrial Electronics, 2017, 65(5): 3752-3762.

[7] Yaman B. Zbede, Shady Gadoue, David Atkinson. Model Predictive MRAS Estimator for Sensorless Induction Motor Drives[J]. IEEE Transactions on Industrial Electronics, 2016, 63(6): 3511-3521.

[8] Tuanjie Li, Jiaxing Zhou. High-Stability Position-Sensorless Control Method for Brushless DC Motors at Low Speed[J]. IEEE Transactions on Power Electronics, 2018, 34(5): 4895-4903.
