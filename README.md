# Open Tracker V1.0

[English](https://www.google.com/search?q=%23english&utm_source=gemini) | [简体中文](https://www.google.com/search?q=%23%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87&utm_source=gemini)

### English

**Open Tracker** is an open-source motion analysis tool implemented in **PyQt6**, **OpenCV**, and **PyQtGraph**, engineered specifically for analyzing massive, high-speed experimental video sequences. By leveraging a dedicated asynchronous `QThread` decoding pipeline with lookahead ring-buffering, it completely eliminates UI latency during playback, rapid keyframe seeking, and dynamic image adjustments.

The platform provides an end-to-end experimental kinematics workflow, featuring two-point physical scale calibration, arbitrary rotated coordinate frames, and cross-ratio projective perspective grids with automatic magnetic snapping. It supports frame-accurate multi-point digitization alongside real-time, interactive displacement time-history plotting, transaction-level undo/redo rollbacks, and automated differential kinematic reporting (component displacements, velocities, and true resultant magnitudes) directly exportable to multi-sheet Excel workbooks.

### 简体中文

**Open Tracker** 是一款基于 **PyQt6**、**OpenCV** 与 **PyQtGraph** 构建的高性能开源运动学分析软件，专为大规模、高帧率的实验视频分析而设计。系统采用独立的后台 `QThread` 解码子线程与预读环形缓存队列，解耦文件 I/O、快速 Seek 定位与图像预处理，从底层消除高分辨率和大文件读取时的 UI 交互卡顿。

软件集成了完整的实验力学数据提取工作流，支持两点真实物理尺度定标、任意倾角的旋转坐标系投影，以及基于射影几何交比定理的透视辅助网格与智能磁吸。它不仅具备帧精确的多目标逐帧数字化捕捉与双向联动的 PyQtGraph 动态位移时程绘制能力，还内置了逐步骤撤回/重做机制，能自动微分计算分向位移、速度与真实合速度，并一键双向交互包含完整实验元数据的结构化 Excel 分析报告。



### License

This project is licensed under the [MIT License](https://www.google.com/search?q=LICENSE&utm_source=gemini).