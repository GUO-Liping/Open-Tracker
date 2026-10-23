# Open Tracker V3.9 — 视频测量、轮廓多点捕捉与点位编辑
# Python 3.9+；安装：python -m pip install PyQt6 pyqtgraph opencv-python numpy pandas openpyxl
# 全屏时两侧默认收起，可通过小三角手动展开；Esc 恢复原布局。
# 双击视频进入全屏；Esc：退出全屏；空格：播放/暂停；P：暂停/继续捕捉
# 暂停捕捉后左键选点/拖动；捕捉时 Ctrl+左键也可修改已有点；右键拖动画面。
# 左侧小三角提示文件树入口；右侧小三角可收起/展开绘图区。
# 文件树显示全机磁盘，默认仅展开当前目录；小三角手动显示/隐藏，右边缘可拖动调宽。
# 每帧支持 1–4000 点；满点自动下一帧，暂停不会丢失已捕捉点。
# 超过 80 点时精简画面编号，点击选择任意点；点数及所选点信息统一显示于状态栏。
# Ctrl+Z 撤销；Ctrl+Y / Ctrl+Shift+Z 还原；单次拖动为一步历史。
# Excel 保留 TrackData / Metadata / ProcessData，并新增可选 PointBasis 表保存逐点标定依据。
# 旧 Excel 无逐点标定依据时，拖动按当前坐标轴/比例换算增量，保留原物理坐标为锚。
# 画面调节支持曝光、伽马、色彩、局部对比度和降噪；工具栏可切换浅色/深色。
# 语言按钮切换全部界面文字；测量支持连续折线、各段长度、总长及顶点内夹角。
# 测量中右键/Enter 结束；时间轴始终为秒，相对首次捕捉帧为零点。
# 视频格式的实际解码能力由安装的 OpenCV/FFmpeg 决定；程序不降低原始分辨率。

import sys
import os
import cv2
import numpy as np
import pandas as pd
import pyqtgraph as pg
import math
import re
import copy
import threading
import time
from bisect import bisect_left
from collections import OrderedDict
from queue import Queue, Empty

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
    QPushButton, QWidget, QLabel, QFileDialog, QInputDialog, QMessageBox, QSlider, QSpinBox,
    QFrame, QSizePolicy, QToolBar, QCheckBox, QComboBox, QToolButton, QMenu,
    QTreeView, QDialog, QFormLayout, QAbstractButton
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QRectF, QPointF, QSize, QThread, QDir, QEvent, QPoint, QModelIndex, QTranslator, QLibraryInfo
from PyQt6.QtGui import (
    QImage, QWheelEvent, QMouseEvent, QPaintEvent, QPainter, QColor, QPen, QFont,
    QAction, QKeySequence, QIcon, QPixmap, QActionGroup, QPolygonF, QFileSystemModel, QPainterPath, QCursor
)

UI_CN_EN = {
'文件':'File', '编辑':'Edit', '分析':'Analysis', '工具':'Tools', '视图':'View', '主工具栏':'Main toolbar',
'加载视频':'Open', '撤回':'Undo', '还原':'Redo', '导出 Excel':'Export', '导入 Excel':'Import',
'导出当前帧':'Save frame', '标定尺寸':'Calibrate', '设置坐标':'Coordinates', '测量查询':'Measure',
'开始捕捉':'Capture', '暂停捕捉':'Pause pts', '继续捕捉':'Resume pts', '透视辅助线':'Guides',
'画面重置':'Reset', '退出':'Exit', '退出视频全屏':'Exit video fullscreen', '画面调节':'Image adjustments',
'浅色界面':'Light', '深色界面':'Dark', '切换为浅色界面':'Switch to light theme', '切换为深色界面':'Switch to dark theme',
'切换为中文':'Switch to Chinese', '切换为英文':'Switch to English',
'光标模式':'Cursor', '屏幕水平/竖直（默认）':'Screen axes (default)', '跟随当前坐标轴':'Follow coordinate axes',
'轨迹模式':'Trails', '无轨迹':'No trails', '局部轨迹 (近10点)':'Recent trails (10 frames)', '全部轨迹':'All trails',
'播放':'Play', '暂停':'Pause', '上一帧':'Previous frame', '下一帧':'Next frame', '帧:':'Frame:', '步长:':'Step:', ' 帧':' frames',
'显示辅助控件':'Show guides', '撤回上一步操作 (Ctrl+Z)':'Undo the last action (Ctrl+Z)',
'还原下一步操作 (Ctrl+Y / Ctrl+Shift+Z)':'Redo the next action (Ctrl+Y / Ctrl+Shift+Z)',
'亮度':'Brightness', '对比度 (%)':'Contrast (%)', '清晰度 (%)':'Sharpness (%)', '曝光 (EV)':'Exposure (EV)',
'伽马':'Gamma', '饱和度 (%)':'Saturation (%)', '色温':'Temperature', '阴影':'Shadows', '高光':'Highlights',
'局部对比度 (%)':'Local contrast (%)', '降噪':'Denoising', '恢复默认':'Reset defaults',
'整体加亮或减暗。':'Brighten or darken the whole image.', '调整灰度差异。':'Adjust tonal contrast.',
'通过反锐化掩模增强边缘细节。':'Enhance edge detail with an unsharp mask.',
'每 +1 EV 将亮度乘以 2；每 −1 EV 减半。':'Each +1 EV doubles brightness; each -1 EV halves it.',
'大于 1 提亮中间调，小于 1 压暗中间调。':'Above 1 brightens midtones; below 1 darkens them.',
'0 为灰度，100 为原始色彩。':'0 gives greyscale; 100 preserves the original saturation.',
'负值偏冷，正值偏暖。':'Negative values cool the image; positive values warm it.',
'主要调整暗部；正值提亮阴影。':'Adjust dark regions; positive values lift shadows.',
'主要调整亮部；负值压低高光。':'Adjust bright regions; negative values reduce highlights.',
'在亮度通道进行自适应对比度增强，辅助观察局部纹理。':'Enhance local luminance contrast to reveal texture.',
'双边滤波，尽量保留边缘；强度增加会增加计算量。':'Edge-preserving bilateral denoising. Higher strengths require more processing.',
'仅调整视频显示；捕捉坐标和原始帧导出保持不变。':'Adjusts the display only. Captured coordinates and raw frame exports are preserved.',
'视频文件树：点击展开/收起，拖动右边缘调整宽度':'Video browser: click to show or hide; drag its right edge to resize',
'拖动调整文件树宽度':'Drag to resize the video browser', '收起右侧绘图区':'Hide the plot panel', '展开右侧绘图区':'Show the plot panel',
'X轴位移 (依自定义坐标系)':'X displacement (custom coordinates)', 'Y轴位移 (依自定义坐标系)':'Y displacement (custom coordinates)',
'{p0}轴位移 (依自定义坐标系)':'{p0} displacement (custom coordinates)', '{p0}轴位移':'{p0} displacement',
'相对时间':'Relative time', '单位':'unit',
'状态: 等待加载视频...':'Status: Waiting for a video...', '状态: 正在后台打开视频…':'Status: Opening video in the background...',
'状态: 加载视频出错':'Status: Video could not be opened', '状态: {p0}':'Status: {p0}',
'已暂停':'Paused', '捕捉中':'Capturing',
'当前帧 {p0}/{p1} 点 | {p2} | 暂停后拖动修改；Ctrl+左键可在捕捉中编辑 | 满点自动下一帧':'Current frame: {p0}/{p1} points | {p2} | Pause to drag points; Ctrl+click to edit while capturing | Auto-advance when complete',
' | 已选 P{p0}（X={p1:.3f}, Y={p2:.3f}）':' | Selected P{p0} (X={p1:.3f}, Y={p2:.3f})',
'空闲。滚轮缩放，右键按住拖动画面，左下角可点击播放/暂停。':'Idle. Scroll to zoom, right-drag to pan, or use Play/Pause below.',
'【尺寸标定】: 请点击定标尺的【起点】 (自动吸附辅助交点/中点)':'Scale calibration: click the start point (snaps to guide intersections and midpoints).',
'【尺寸标定】: 请点击定标尺的【终点】 (自动吸附辅助交点/中点)':'Scale calibration: click the end point (snaps to guide intersections and midpoints).',
'【建立坐标系】: 请点击设定【坐标原点 (0,0)】 (自动吸附辅助点)':'Coordinates: click the origin (0,0), with snapping to guide points.',
'【建立坐标系】: 请点击画面确定【X轴正方向】(随后会弹窗供你微调角度)':'Coordinates: click towards positive X; the angle can then be refined.',
'【测量/查询】: 连续测量模式。请点击起点测量距离 (标定后显示真实物理尺寸)。':'Measurement: click a start point. Calibrate the scale to obtain physical distances.',
'【测量/查询】: 请点击终点以固定测量线段，随后可继续下一次测量。':'Measurement: click the end point, then continue with the next measurement.',
'【捕捉模式】: 设置为每帧 {p0} 个点。请点击目标（当前第 {p1}/{p2} 个点）':'Capture: {p0} points per frame. Click point {p1}/{p2}.',
'【捕捉已暂停】左键选中并拖动当前帧的点；P 继续。':'Capture paused. Select and drag a point in this frame; press P to resume.',
'【辅助线模式】: 请画出透视视场下的 4 条轮廓线段，将利用交比自动计算视场中心和透视中点。':'Perspective guides: draw four boundary lines to calculate the centre and projective midpoints.',
'开启捕捉模式':'Capture settings', '每帧目标点数（按轮廓顺序点击，支持数百点）：':'Points per frame (click in contour order; hundreds of points are supported):',
'尚未设置坐标':'Coordinates not set', '请先设置坐标原点和 X 轴方向。':'Set the coordinate origin and positive X direction first.',
'标定物理尺寸':'Calibrate physical scale', '像素跨度: {p0:.2f} px\n请输入实际物理长度:':'Pixel distance: {p0:.2f} px\nEnter the physical length:',
'设置单位名称':'Set unit label', '请输入单位简称 (如 mm, m):':'Enter the unit symbol (e.g. mm or m):',
'标定映射建立！\n1 像素 = {p0:.4f} {p1}':'Scale calibrated.\n1 pixel = {p0:.4f} {p1}',
'精调坐标轴角度':'Refine coordinate angle',
'你已点击定义了X轴方向。\n在此微调逆时针旋转角度(度)或直接确认\n(0 为默认水平向右):':'The positive X direction has been selected.\nRefine the anticlockwise angle in degrees, or confirm.\n(0 is horizontal, towards the right):',
'设置成功':'Coordinates set', '自定义坐标系已生效！\n原点: ({p0:.1f}, {p1:.1f})\nX轴倾角: {p2:.2f}°':'Custom coordinates are active.\nOrigin: ({p0:.1f}, {p1:.1f})\nX-axis angle: {p2:.2f}°',
'选择视频文件':'Choose a video', '视频打开错误':'Video error', 'OpenCV 无法打开该视频，请检查文件及编码格式。':'OpenCV could not open this video. Check the file and codec.',
'切换视频':'Switch video', '切换视频将清除当前捕捉记录及撤销历史。请先导出需要保留的数据。继续打开？':'Opening another video clears the current captures and undo history. Export any data you need to keep. Continue?',
'无法导出':'Cannot export', '请先加载视频并定位到需要导出的帧。':'Open a video and navigate to the frame you want to export.',
'导出当前帧原始画面':'Export the original frame', '无法从视频中读取当前帧。':'Could not read the current frame from the video.',
'成功':'Success', '导出失败':'Export failed', '写入图像文件失败，请检查文件路径或格式。':'Could not save the image. Check the path and format.',
'当前帧原始画面已导出：\n{p0}':'The original frame has been exported:\n{p0}',
'数据存档为 Excel':'Save capture data as Excel', '报告成功写入到本地：\n{p0}':'The report has been saved:\n{p0}',
'导入 Excel 捕捉数据':'Import Excel capture data', '导入失败':'Import failed', '无法识别帧索引列 (Frame_Index)':'The frame index column (Frame_Index) could not be identified.',
'视频不匹配':'Video mismatch', '导入数据来自:\n{p0}\n\n当前加载的是:\n{p1}\n\n仍要恢复标定参数并导入数据吗？':'Imported data belongs to:\n{p0}\n\nThe current video is:\n{p1}\n\nRestore the calibration and import the data anyway?',
'导入成功':'Import complete', '成功导入 {p0} 帧记录 (每帧 {p1} 点)':'Imported {p0} frames ({p1} points per frame)',
'\n（跳过了 {p0} 条超出当前视频帧范围的记录）':'\n(Skipped {p0} records outside the current video frame range.)',
'读取 Excel 文件出错：\n{p0}':'Could not read the Excel file:\n{p0}',
'✌️视频地址 {p0}\n🖼️分辨率 {p1}  |  🎬 帧率 {p2} FPS  |  📊 总帧数 {p3}  |  ⏱️ 时长 {p4}':'Video: {p0}\nResolution: {p1}  |  Frame rate: {p2} FPS  |  Frames: {p3}  |  Duration: {p4}',
'📍 当前坐标: {p0}':'Coordinates: {p0}', '📏 实时长度: {p0}':'Live distance: {p0}',
'📏 测量距离: {p0}\n🔸 起点: {p1}\n🔹 终点: {p2}':'Distance: {p0}\nStart: {p1}\nEnd: {p2}',
'新增 P{p0}':'Add P{p0}', '移动 P{p0}':'Move P{p0}', '画面点击: {p0}':'Canvas click: {p0}', '进入 {p0}':'Enter {p0}',
'绘制屏幕尺寸固定的小十字，缩放画面时仍保持简洁清晰。':'Draw a fixed-size screen cross.',
'在捕捉点旁绘制与对应轨迹同色的圆形数字标记。':'Draw a numbered marker beside a captured point.',
}


UI_CN_EN.update({'开始折线测量': 'Start polyline measurement', '结束折线测量': 'Finish polyline measurement', '折线测量：逐点点击；右键或 Enter 结束本次测量。': 'Polyline measurement: click successive points; right-click or press Enter to finish.', '测量：{p0} 点 | 总长 {p1:.2f} {p2}': 'Measurement: {p0} points | Total {p1:.2f} {p2}', '各段': 'Segments', '内夹角': 'Interior angles', '已结束；左键开始下一次测量。': 'Finished; click to start the next measurement.', '右键或 Enter 结束。': 'Right-click or press Enter to finish.', '总长: {p0:.2f} {p1} | {p2} 点': 'Total: {p0:.2f} {p1} | {p2} points', '测量已结束；左键开始下一次。': 'Measurement finished; click to start another.', '右键或 Enter 结束本次测量。': 'Right-click or press Enter to finish this measurement.', '预览: {p0:.2f} {p1}': 'Preview: {p0:.2f} {p1}', '绘图选项': 'Plot Options', '变换': 'Transforms', '降采样': 'Downsample', '平均': 'Average', '透明度': 'Alpha', '网格': 'Grid', '数据点': 'Points', '不透明度': 'Opacity', 'X 对数轴': 'Log X', 'Y 对数轴': 'Log Y', '求导 dy/dx': 'dy/dx', 'Y 与 Y′': "Y vs. Y'", '功率谱 (FFT)': 'Power Spectrum (FFT)', '减去均值': 'Subtract Mean', '裁剪到视野': 'Clip to View', '最大曲线数:': 'Max Traces:', '峰值': 'Peak', '忘记隐藏曲线': 'Forget hidden traces', '均值': 'Mean', '抽取': 'Subsample', '自动': 'Auto', '显示 X 网格': 'Show X Grid', '显示 Y 网格': 'Show Y Grid', '视图选项': 'ViewBox options', '显示全部': 'View All', 'X 轴': 'X axis', 'Y 轴': 'Y axis', '鼠标模式': 'Mouse Mode', '三键模式': '3 button', '单键模式': '1 button', '关联坐标轴:': 'Link Axis:', '手动': 'Manual', '反转坐标轴': 'Invert Axis', '启用鼠标': 'Mouse Enabled', '仅可见数据': 'Visible Data Only', '仅自动平移': 'Auto Pan Only'})
UI_EN_CN = {english: chinese for chinese, english in UI_CN_EN.items()}
CURRENT_LANGUAGE = 'zh'

def ui(text):
    if not isinstance(text, str):
        return text
    if CURRENT_LANGUAGE == 'en':
        return UI_CN_EN.get(text, text)
    return UI_EN_CN.get(text, text)

def ui_format(template, **values):
    return ui(template).format(**values)

VIDEO_EXTENSIONS = ('.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.mpg', '.mpeg',
                    '.webm', '.mts', '.m2ts', '.ts', '.vob', '.3gp', '.ogv', '.mxf', '.asf')

cv2.setUseOptimized(True)
pg.setConfigOptions(background="#101418", foreground="#d7dde8", antialias=True)

# 轨迹曲线颜色库
TRACK_COLORS_X = ["#5ac8fa", "#ff9800", "#8cc63f", "#00ffff", "#c678dd", "#ffeb3b", "#00bcd4", "#4caf50"]
TRACK_COLORS_Y = ["#ff5c8a", "#ff4d4f", "#f7931e", "#9e005d", "#3f51b5", "#e91e63", "#00ff00", "#ffffff"]
TRACK_COLORS_VIEW = ["#ff4d4f", "#5ac8fa", "#ffeb3b", "#8cc63f", "#c678dd", "#f7931e", "#00ffff", "#ff5c8a"]

APP_STYLE = """
QMainWindow, QWidget { background: #0f141b; color: #d7dde8; font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif; font-size: 10pt; }
QMenuBar { background: #111821; color: #d7dde8; border-bottom: 1px solid #263241; }
QMenuBar::item { padding: 6px 10px; background: transparent; }
QMenuBar::item:selected { background: #202a36; }
QMenu { background: #151c25; border: 1px solid #344255; color: #d7dde8; }
QMenu::item { padding: 7px 28px; }
QMenu::item:selected { background: #263241; }
QToolBar { background: #151c25; border: 0; border-bottom: 1px solid #263241; spacing: 4px; padding: 5px; }
QToolButton { background: transparent; border: 1px solid transparent; border-radius: 6px; padding: 5px 8px; color: #d7dde8; }
QToolButton:hover { background: #202a36; border-color: #344255; }
QToolButton:pressed { background: #111821; }
QToolButton:disabled { color: #687484; }
QToolButton::menu-indicator { image: none; width: 0px; }
QFrame#controlBar, QFrame#statusBar { background: #151c25; border: 1px solid #263241; border-radius: 8px; }
QPushButton { background: #202a36; border: 1px solid #344255; border-radius: 6px; color: #edf2f8; padding: 7px 12px; min-height: 24px; }
QPushButton::icon { color: #edf2f8; }
QPushButton:hover { background: #2a3747; border-color: #4c5e74; }
QPushButton:pressed { background: #18212c; }
QPushButton:disabled { color: #687484; background: #151b23; border-color: #202936; }
QPushButton#primaryButton { background: #1267d3; border-color: #3181ee; font-weight: 600; }
QPushButton#primaryButton:hover { background: #1976ed; }
QSlider::groove:horizontal { height: 6px; background: #263241; border-radius: 3px; }
QSlider::handle:horizontal { width: 16px; margin: -5px 0; border-radius: 8px; background: #5aa7ff; }
QSpinBox { background: #111821; border: 1px solid #344255; border-radius: 6px; padding: 5px 28px 5px 8px; min-height: 28px; }
QSpinBox::up-button, QSpinBox::down-button { width: 26px; border-left: 1px solid #344255; background: #202a36; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover { background: #2a3747; }
QComboBox { background: #111821; border: 1px solid #344255; border-radius: 6px; padding: 4px 10px; color: #d7dde8; min-height: 28px; }
QComboBox::drop-down { border: 0; }
QComboBox QAbstractItemView { background: #151c25; border: 1px solid #344255; color: #d7dde8; selection-background-color: #263241; }
QCheckBox { spacing: 6px; }
QCheckBox::indicator { width: 16px; height: 16px; }
QLabel#statusLabel { color: #ffd166; font-weight: 600; }
"""


# 新增显示调节项：名称、标签、滑块范围、初值、数值换算、说明。
EXTRA_ADJUSTMENT_SPECS = (
    ('exposure', '曝光 (EV)', -200, 200, 0, 100.0, '每 +1 EV 将亮度乘以 2；每 −1 EV 减半。'),
    ('gamma', '伽马', 30, 300, 100, 100.0, '大于 1 提亮中间调，小于 1 压暗中间调。'),
    ('saturation', '饱和度 (%)', 0, 200, 100, 100.0, '0 为灰度，100 为原始色彩。'),
    ('temperature', '色温', -100, 100, 0, 100.0, '负值偏冷，正值偏暖。'),
    ('shadows', '阴影', -100, 100, 0, 100.0, '主要调整暗部；正值提亮阴影。'),
    ('highlights', '高光', -100, 100, 0, 100.0, '主要调整亮部；负值压低高光。'),
    ('local_contrast', '局部对比度 (%)', 0, 100, 0, 100.0, '在亮度通道进行自适应对比度增强，辅助观察局部纹理。'),
    ('denoise', '降噪', 0, 10, 0, 1.0, '双边滤波，尽量保留边缘；强度增加会增加计算量。'),
)
EXTRA_DEFAULTS = {name: initial/divisor for name, _, _, _, initial, divisor, _ in EXTRA_ADJUSTMENT_SPECS}

def _light_style():
    colors = {
        '#0f141b': '#f4f6fa', '#d7dde8': '#263447', '#111821': '#ffffff',
        '#263241': '#d7dfeb', '#202a36': '#edf2f8', '#151c25': '#ffffff',
        '#344255': '#c6d1df', '#edf2f8': '#223047', '#2a3747': '#e2ebf7',
        '#4c5e74': '#8da9ce', '#18212c': '#d9e5f5', '#687484': '#9aa6b5',
        '#151b23': '#eef1f5', '#202936': '#dde3eb', '#ffd166': '#866100'}
    return re.sub(r'#[0-9a-fA-F]{6}', lambda m: colors.get(m.group().lower(), m.group()), APP_STYLE)

LIGHT_STYLE = _light_style()

# ==========================================================
# 1. 多线程后台解码工作线程 (VideoDecoderWorker)
# ==========================================================

class VideoDecoderWorker(QThread):
    """仅此线程访问 VideoCapture；请求带版本号，旧帧不会覆盖新视频或定位结果。"""
    frame_loaded = pyqtSignal(int, QImage, object, int)
    video_info_loaded = pyqtSignal(dict)
    error_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.cap = None
        self.buffer_queue = Queue(maxsize=4)
        self._condition = threading.Condition()
        self.generation = 0
        self._path = None
        self._seek = None
        self.is_playing = False
        self.is_running = True
        self.play_next_idx = 0
        self.play_step = 1
        self.adjustments = dict(EXTRA_DEFAULTS, brightness=0, contrast=1.0, sharpness=0.0)
        self.decoder_next_idx = -1
        self._last_idx = -1
        self._last_raw = None

    def clear_buffer(self):
        while True:
            try:
                self.buffer_queue.get_nowait()
            except Empty:
                break

    def _invalidate(self):
        self.generation += 1
        self.clear_buffer()
        self._condition.notify_all()

    def open_video(self, path):
        with self._condition:
            self._path = path
            self._seek = 0
            self.is_playing = False
            self.adjustments = dict(EXTRA_DEFAULTS, brightness=0, contrast=1.0, sharpness=0.0)
            self._invalidate()

    def request_seek(self, index):
        with self._condition:
            self.is_playing = False
            self._seek = int(index)
            self._invalidate()

    def start_playback(self, start_idx, step=1):
        with self._condition:
            self._seek = None
            self.play_next_idx = int(start_idx)
            self.play_step = max(1, int(step))
            self.is_playing = True
            self._invalidate()

    def stop_playback(self):
        with self._condition:
            self.is_playing = False
            self._invalidate()

    def set_adjustments(self, brightness, contrast, sharpness=0.0, refresh_idx=None, extra_adjustments=None):
        with self._condition:
            value = dict(EXTRA_DEFAULTS)
            value.update(extra_adjustments or {})
            value.update(brightness=int(brightness), contrast=float(contrast), sharpness=float(sharpness))
            if value == self.adjustments:
                return
            self.adjustments = value
            if refresh_idx is not None and refresh_idx >= 0:
                self._seek = int(refresh_idx)
                if self.is_playing:
                    self.play_next_idx = int(refresh_idx) + self.play_step
            self._invalidate()

    def run(self):
        while True:
            with self._condition:
                if not self.is_running:
                    break
                path = self._path
                self._path = None
                gen = self.generation
                target = self._seek
                self._seek = None
                seek_job = target is not None
                if target is None and self.is_playing and not self.buffer_queue.full():
                    target = self.play_next_idx
                settings = self.adjustments
                if path is None and target is None:
                    self._condition.wait(0.01)
                    continue
            if path is not None:
                if self.cap is not None:
                    self.cap.release()
                self.cap = cv2.VideoCapture(path)
                self.decoder_next_idx = -1
                self._last_idx, self._last_raw = -1, None
                if not self.cap.isOpened():
                    with self._condition:
                        current = gen == self.generation
                    if current:
                        self.error_signal.emit('OpenCV 无法打开该视频，请检查文件及编码格式。')
                    continue
                fps = self.cap.get(cv2.CAP_PROP_FPS)
                fps = fps if math.isfinite(fps) and fps > 0 else 30.0
                with self._condition:
                    if gen != self.generation:
                        continue
                    self.video_info_loaded.emit({
                        'total_frames': max(1, int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))),
                        'fps': fps, 'width': int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)),
                        'height': int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)),
                        'path': path, 'generation': gen})
            if target is None or self.cap is None or not self.cap.isOpened():
                continue
            raw = self._read_frame_at(target)
            if raw is None:
                with self._condition:
                    if gen == self.generation:
                        self.is_playing = False
                continue
            qimg, raw = self._process_frame(raw, settings)
            with self._condition:
                if gen != self.generation:
                    continue
                if seek_job:
                    self.frame_loaded.emit(target, qimg, raw, gen)
                elif self.is_playing:
                    self.buffer_queue.put_nowait((target, qimg, raw, gen))
                    self.play_next_idx = target + self.play_step
        if self.cap is not None:
            self.cap.release()

    def _read_frame_at(self, target_idx):
        if target_idx == self._last_idx and self._last_raw is not None:
            return self._last_raw
        diff = target_idx - self.decoder_next_idx
        if self.decoder_next_idx >= 0 and 0 <= diff <= 9:
            for _ in range(diff):
                self.cap.grab()
        else:
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, target_idx)
        ok, frame = self.cap.read()
        if not ok:
            return None
        self.decoder_next_idx = target_idx + 1
        self._last_idx, self._last_raw = target_idx, frame
        return frame

    def _process_frame(self, frame, settings=None):
        config = settings if settings is not None else self.adjustments
        if isinstance(config, (tuple, list)):  # 兼容原来的三参数调用。
            config = dict(EXTRA_DEFAULTS, brightness=config[0], contrast=config[1], sharpness=config[2])
        values = dict(EXTRA_DEFAULTS, brightness=0, contrast=1.0, sharpness=0.0)
        values.update(config)
        out = frame
        tone_names = ('brightness', 'contrast', 'exposure', 'gamma', 'shadows', 'highlights')
        defaults = dict(EXTRA_DEFAULTS, brightness=0, contrast=1.0)
        if any(abs(values[name]-defaults[name]) > 1e-6 for name in tone_names):
            level = np.arange(256, dtype=np.float32)/255.0
            level = np.clip(level * (2.0 ** values['exposure']), 0, 1)
            level = np.power(level, 1.0/max(values['gamma'], 0.01))
            level = level + 0.3*values['shadows']*(1-level)**2 + 0.3*values['highlights']*level**2
            lut = np.clip(level*255*values['contrast']+values['brightness'], 0, 255).astype(np.uint8)
            out = cv2.LUT(out, lut)
        temperature = values['temperature']
        if abs(temperature) > 1e-6:
            channels = cv2.split(out)
            factors = (1-0.18*temperature, 1.0, 1+0.18*temperature)
            out = cv2.merge([cv2.LUT(channel, np.clip(np.arange(256)*factor, 0, 255).astype(np.uint8))
                             for channel, factor in zip(channels, factors)])
        if abs(values['saturation']-1.0) > 1e-6:
            hsv = cv2.cvtColor(out, cv2.COLOR_BGR2HSV)
            lut = np.clip(np.arange(256)*values['saturation'], 0, 255).astype(np.uint8)
            hsv[:, :, 1] = cv2.LUT(hsv[:, :, 1], lut)
            out = cv2.cvtColor(hsv, cv2.COLOR_HSV2BGR)
        if values['local_contrast'] > 0:
            lab = cv2.cvtColor(out, cv2.COLOR_BGR2LAB)
            strength = float(values['local_contrast'])
            if getattr(self, '_clahe_strength', None) != strength:
                self._clahe_strength = strength
                self._clahe = cv2.createCLAHE(clipLimit=1+3*strength, tileGridSize=(8, 8))
            enhanced = self._clahe.apply(lab[:, :, 0])
            lab[:, :, 0] = cv2.addWeighted(lab[:, :, 0], 1-strength, enhanced, strength, 0)
            out = cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)
        if values['denoise'] > 0:
            out = cv2.bilateralFilter(out, 5, values['denoise']*5, 5)
        if values['sharpness'] > 0:
            blur = cv2.GaussianBlur(out, (0, 0), 1.0)
            out = cv2.addWeighted(out, 1+values['sharpness'], blur, -values['sharpness'], 0)
        rgb = cv2.cvtColor(out, cv2.COLOR_BGR2RGB)
        qimg = QImage(rgb.data, rgb.shape[1], rgb.shape[0], rgb.strides[0], QImage.Format.Format_RGB888).copy()
        return qimg, frame

    def stop(self):
        with self._condition:
            self.is_running = False
            self._condition.notify_all()
        self.wait()


# ==========================================================
# 2. 视频画布渲染组件 (VideoWidget)
# ==========================================================

class VideoWidget(QWidget):
    clicked_pos = pyqtSignal(float, float)
    moved_pos = pyqtSignal(float, float)
    zoom_changed = pyqtSignal(int)
    point_pressed = pyqtSignal(int)
    point_dragged = pyqtSignal(float, float)
    point_released = pyqtSignal()
    fullscreen_requested = pyqtSignal(bool)
    measurement_finished = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.canvas_background = QColor("#05070a")
        self.setStyleSheet("background-color: #05070a; border: 1px solid #263241; border-radius: 8px;")
        self.setCursor(Qt.CursorShape.CrossCursor)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumSize(320, 240)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.zoom_level = 1.0
        self.pan_offset_x = 0.0
        self.pan_offset_y = 0.0
        
        self.is_panning = False
        self.last_mouse_pos = None
        self.canvas_w = 0
        self.canvas_h = 0
        self.frame_qimage = None
        self.overlay_state = {}
        self._mark_layer_key = None
        self._mark_layer = None
        self._dragging_point = False
        self._left_down_pos = None
        self._drag_started = False
        self._last_left_was_click = False
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent, True)

    def set_frame(self, qimg):
        self.frame_qimage = qimg
        self.update()

    def set_overlay_state(self, **state):
        self.overlay_state = state
        self.update()

    def fit_to_view(self):
        if self.canvas_w <= 0 or self.canvas_h <= 0 or self.width() <= 0: return
        zw = self.width() / self.canvas_w
        zh = self.height() / self.canvas_h
        self.zoom_level = min(zw, zh) * 0.98
        self.pan_offset_x = 0.0
        self.pan_offset_y = 0.0
        self.zoom_changed.emit(int(self.zoom_level * 100))
        self.update()

    def paintEvent(self, event: QPaintEvent):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, self.zoom_level < 3.0)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.fillRect(self.rect(), self.canvas_background)

        if self.frame_qimage:
            view_rect = self._view_rect()
            painter.drawImage(view_rect, self.frame_qimage, QRectF(self.frame_qimage.rect()))
            self._paint_overlays(painter, view_rect)
        painter.end()

    def reset_view(self):
        self.zoom_level = 1.0
        self.pan_offset_x = 0.0
        self.pan_offset_y = 0.0
        self.zoom_changed.emit(int(self.zoom_level * 100))
        self.update()

    def _view_rect(self):
        view_w = self.canvas_w * self.zoom_level
        view_h = self.canvas_h * self.zoom_level
        x_offset = (self.width() - view_w) / 2 + self.pan_offset_x
        y_offset = (self.height() - view_h) / 2 + self.pan_offset_y
        return QRectF(x_offset, y_offset, view_w, view_h)

    def _get_canvas_coords(self, pos):
        if self.canvas_w <= 0 or self.canvas_h <= 0: return None
        view_rect = self._view_rect()
        x_offset = view_rect.x()
        y_offset = view_rect.y()
        canvas_x = (pos.x() - x_offset) / self.zoom_level
        canvas_y = (pos.y() - y_offset) / self.zoom_level
        if 0 <= canvas_x < self.canvas_w and 0 <= canvas_y < self.canvas_h:
            return float(canvas_x), float(canvas_y)
        return None

    def _to_screen_point(self, canvas_pt):
        x, y = canvas_pt
        rect = self._view_rect()
        return QPointF(rect.x() + x * self.zoom_level, rect.y() + y * self.zoom_level)

    def _paint_overlays(self, painter, view_rect):
        state = self.overlay_state
        if not state: return
        widget_rect = self.rect()
        painter.setFont(QFont("Segoe UI", 10))

        def draw_textbox(text, x, y, bg_color=QColor(0,0,0,180)):
            painter.setFont(QFont("Segoe UI", 9))
            fm = painter.fontMetrics()
            lines = text.split('\n')
            text_w = max(fm.horizontalAdvance(ln) for ln in lines) + 24
            text_h = fm.lineSpacing() * len(lines) + 16
            rect = QRectF(x, y, text_w, text_h)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(bg_color)
            painter.drawRoundedRect(rect, 6, 6)
            painter.setPen(QColor("#edf2f8"))
            painter.drawText(rect.adjusted(12, 8, -12, -8), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, ui(text))

        def draw_point(pt, color, radius=4):
            screen = self._to_screen_point(pt)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(color))
            painter.drawEllipse(screen, radius, radius)
            return screen

        def draw_cross(pt, color, size=5, width=1.2):
            ui("""绘制屏幕尺寸固定的小十字，缩放画面时仍保持简洁清晰。""")
            screen = self._to_screen_point(pt)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.setPen(QPen(QColor(color), width, Qt.PenStyle.SolidLine,
                                Qt.PenCapStyle.RoundCap))
            painter.drawLine(QPointF(screen.x() - size, screen.y()),
                             QPointF(screen.x() + size, screen.y()))
            painter.drawLine(QPointF(screen.x(), screen.y() - size),
                             QPointF(screen.x(), screen.y() + size))
            return screen

        def draw_number_badge(pt, text, color, offset_x=-12):
            ui("""在捕捉点旁绘制与对应轨迹同色的圆形数字标记。""")
            anchor = self._to_screen_point(pt)
            center = anchor + QPointF(offset_x, -12)
            radius = 8.0
            badge_color = QColor(color)
            painter.setPen(QPen(badge_color, 1.5, Qt.PenStyle.SolidLine))
            painter.setBrush(QColor(255, 255, 255, 0))
            painter.drawEllipse(center, radius, radius)
            painter.setPen(badge_color)
            painter.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            painter.drawText(
                QRectF(center.x() - radius, center.y() - radius,
                       radius * 2, radius * 2),
                Qt.AlignmentFlag.AlignCenter,
                ui(text)
            )

        scale_factor = state.get("scale_factor", 1.0)
        physics_unit = state.get("physics_unit", ui("单位"))
        is_calibrated = (state.get("calib_line") is not None)
        coord_origin = state.get("coord_origin")
        coord_angle = state.get("coord_angle", 0.0)
        show_overlays = state.get("show_overlays", True)
        cursor_mode = state.get("cursor_mode", 0)
        mode = state.get("mode")
        mouse_pos = state.get("mouse_curr_pos")

        def get_formatted_coords(pt):
            if is_calibrated and coord_origin:
                dx = pt[0] - coord_origin[0]
                dy = coord_origin[1] - pt[1]
                rx = (dx * math.cos(coord_angle) + dy * math.sin(coord_angle)) * scale_factor
                ry = (-dx * math.sin(coord_angle) + dy * math.cos(coord_angle)) * scale_factor
                return f"X:{rx:.1f} Y:{ry:.1f} {physics_unit}"
            elif is_calibrated:
                return f"X:{pt[0]*scale_factor:.1f} Y:{pt[1]*scale_factor:.1f} {physics_unit}"
            else:
                return f"X:{int(pt[0])} Y:{int(pt[1])} px"

        video_info = state.get("video_info")
        if video_info:
            display_path = video_info['path']
            if len(display_path) > 100: display_path = display_path[:30] + " ... " + display_path[-90:]
            info_text = (
                ui_format('✌️视频地址 {p0}\n🖼️分辨率 {p1}  |  🎬 帧率 {p2} FPS  |  📊 总帧数 {p3}  |  ⏱️ 时长 {p4}', p0=display_path, p1=video_info['resolution'], p2=video_info['fps'], p3=video_info['frames'], p4=video_info['duration'])
            )
            draw_textbox(info_text, 8, 8)

        # =============== 控件图层隔离 ===============
        if show_overlays:
            aux_lines = state.get("aux_lines", [])
            aux_current = state.get("aux_current_points", [])
            aux_intersections = state.get("aux_intersections", [])
            aux_midpoints = state.get("aux_midpoints", [])
            aux_complete = len(aux_intersections) == 4
            
            # 绘制阶段显示轮廓线和端点；网格完成后隐藏这些临时标记。
            if not aux_complete:
                painter.setPen(QPen(QColor(0, 255, 255, 150), 1.5, Qt.PenStyle.SolidLine))
                for line in aux_lines:
                    painter.drawLine(self._to_screen_point(line[0]), self._to_screen_point(line[1]))
                    draw_cross(line[0], "#00ffff", 3)
                    draw_cross(line[1], "#00ffff", 3)
                
            if aux_current and len(aux_current) == 1:
                draw_cross(aux_current[0], "#00ffff", 3)
                if state.get("mouse_curr_pos"):
                    painter.setPen(QPen(QColor(0, 255, 255, 180), 1.5, Qt.PenStyle.DashLine))
                    painter.drawLine(self._to_screen_point(aux_current[0]), self._to_screen_point(state["mouse_curr_pos"]))
            
            if len(aux_intersections) == 4:
                C0, C1, C2, C3 = aux_intersections
                painter.setPen(QPen(QColor(0, 255, 255, 100), 1.0, Qt.PenStyle.SolidLine))
                p0, p1, p2, p3 = self._to_screen_point(C0), self._to_screen_point(C1), self._to_screen_point(C2), self._to_screen_point(C3)
                painter.drawLine(p0, p1); painter.drawLine(p1, p2)
                painter.drawLine(p2, p3); painter.drawLine(p3, p0)
                
                grid_size = state.get("grid_size", 10.0)
                if grid_size <= 0: grid_size = 10.0
                
                len_top = math.hypot(C1[0]-C0[0], C1[1]-C0[1])
                len_left = math.hypot(C3[0]-C0[0], C3[1]-C0[1])
                phys_w = len_top * scale_factor
                phys_h = len_left * scale_factor
                steps_x = min(200, max(1, int(round(phys_w / grid_size))))
                steps_y = min(200, max(1, int(round(phys_h / grid_size))))
                
                painter.setPen(QPen(QColor(0, 255, 255, 60), 1.0, Qt.PenStyle.DashLine))
                for i in range(1, steps_x):
                    t = i / steps_x
                    pt_top_x = C0[0] + t * (C1[0] - C0[0]); pt_top_y = C0[1] + t * (C1[1] - C0[1])
                    pt_bot_x = C3[0] + t * (C2[0] - C3[0]); pt_bot_y = C3[1] + t * (C2[1] - C3[1])
                    painter.drawLine(self._to_screen_point((pt_top_x, pt_top_y)), self._to_screen_point((pt_bot_x, pt_bot_y)))
                for i in range(1, steps_y):
                    t = i / steps_y
                    pt_left_x = C0[0] + t * (C3[0] - C0[0]); pt_left_y = C0[1] + t * (C3[1] - C0[1])
                    pt_right_x = C1[0] + t * (C2[0] - C1[0]); pt_right_y = C1[1] + t * (C2[1] - C1[1])
                    painter.drawLine(self._to_screen_point((pt_left_x, pt_left_y)), self._to_screen_point((pt_right_x, pt_right_y)))

            # 完成后仅保留透明网格及极简十字，不再显示交点/中点/中心文字。
            if aux_complete:
                for pt in aux_intersections:
                    draw_cross(pt, QColor(0, 255, 255, 135), 3.5, 1.0)
                for pt in aux_midpoints:
                    draw_cross(pt, QColor(229, 192, 123, 135), 3.0, 1.0)

            # 标定主线
            if state.get("calib_pt1"):
                p1 = self._to_screen_point(state.get("calib_pt1"))
                draw_point(state.get("calib_pt1"), "#ff4d4f", 4)
                if mode == "CALIBRATING_2" and mouse_pos:
                    p2 = self._to_screen_point(mouse_pos)
                    painter.setPen(QPen(QColor("#ffd43b"), 1.3))
                    painter.drawLine(p1, p2)

            calib_line = state.get("calib_line")
            if calib_line:
                p1 = self._to_screen_point(calib_line["p1"])
                p2 = self._to_screen_point(calib_line["p2"])
                painter.setPen(QPen(QColor("#ffd43b"), 1.5))
                painter.drawLine(p1, p2)
                draw_point(calib_line["p1"], "#ffd43b", 4); draw_point(calib_line["p2"], "#ffd43b", 4)
                painter.setPen(QPen(QColor("#ffd43b")))
                mid = QPointF((p1.x() + p2.x()) / 2, (p1.y() + p2.y()) / 2)
                painter.drawText(mid + QPointF(8, -8), ui(calib_line["text"]))

            # 自定义坐标系 
            if mode == "SET_COORD_2" and coord_origin and mouse_pos:
                painter.setPen(QPen(QColor(255, 64, 64, 210), 1.0, Qt.PenStyle.SolidLine))
                painter.drawLine(self._to_screen_point(coord_origin), self._to_screen_point(mouse_pos))
                
            if coord_origin:
                origin_scr = self._to_screen_point(coord_origin)
                L = max(widget_rect.width(), widget_rect.height()) * 2 
                ux, uy = math.cos(coord_angle), -math.sin(coord_angle)
                vx, vy = -math.sin(coord_angle), -math.cos(coord_angle)

                axis_pen = QPen(QColor(255, 64, 64, 190), 1.0, Qt.PenStyle.SolidLine)
                painter.setPen(axis_pen)
                painter.drawLine(self._to_screen_point((coord_origin[0] - L*ux, coord_origin[1] - L*uy)), 
                                 self._to_screen_point((coord_origin[0] + L*ux, coord_origin[1] + L*uy)))
                painter.setPen(axis_pen)
                painter.drawLine(self._to_screen_point((coord_origin[0] - L*vx, coord_origin[1] - L*vy)), 
                                 self._to_screen_point((coord_origin[0] + L*vx, coord_origin[1] + L*vy)))

                painter.setPen(QPen(QColor("#ff4040")))
                painter.drawText(self._to_screen_point((coord_origin[0] + 120*ux, coord_origin[1] + 120*uy)) + QPointF(5, -5), ui("X"))
                painter.setPen(QPen(QColor("#ff4040")))
                painter.drawText(self._to_screen_point((coord_origin[0] + 120*vx, coord_origin[1] + 120*vy)) + QPointF(5, -5), ui("Y"))
                draw_cross(coord_origin, "#ff4040", 4, 1.0)

            # 连续折线测量：各段长度、各顶点内夹角，以及总长度。
            points = state.get('measure_points', [])
            data = state.get('measure_line') or {}
            segments, angles = data.get('segments', []), data.get('angles', [])
            unit = data.get('unit', 'px')
            for i, (a,b) in enumerate(zip(points,points[1:])):
                sa,sb = self._to_screen_point(a),self._to_screen_point(b)
                painter.setPen(QPen(QColor('#00bcd4'),2))
                painter.drawLine(sa,sb)
                mid = (sa+sb)*0.5
                draw_textbox(f'L{i+1}: {segments[i]:.2f} {unit}',mid.x()+6,mid.y()+6,QColor(0,70,80,205))
            for i, point in enumerate(points):
                scr = draw_point(point,'#00d9ef',4)
                painter.setPen(QColor('#00d9ef'))
                painter.drawText(scr+QPointF(6,-8),f'P{i+1}')
                if 0<i<len(points)-1 and angles[i-1] is not None:
                    a,b = self._to_screen_point(points[i-1])-scr,self._to_screen_point(points[i+1])-scr
                    start = math.degrees(math.atan2(-a.y(),a.x()))
                    end = math.degrees(math.atan2(-b.y(),b.x()))
                    span = (end-start+180)%360-180
                    painter.setBrush(Qt.BrushStyle.NoBrush)
                    painter.setPen(QPen(QColor('#ffd166'),1.5))
                    painter.drawArc(QRectF(scr.x()-18,scr.y()-18,36,36),int(start*16),int(span*16))
                    draw_textbox(f'{angles[i-1]:.2f}°',scr.x()+8,scr.y()-42,QColor(80,60,0,205))
            if points and mode == 'MEASURE_2' and mouse_pos:
                last,mouse = self._to_screen_point(points[-1]),self._to_screen_point(mouse_pos)
                painter.setPen(QPen(QColor('#00bcd4'),1.4,Qt.PenStyle.DashLine))
                painter.drawLine(last,mouse)
                distance = math.hypot(mouse_pos[0]-points[-1][0],mouse_pos[1]-points[-1][1])
                distance *= scale_factor if is_calibrated else 1.0
                mid = (last+mouse)*0.5
                draw_textbox(ui_format('预览: {p0:.2f} {p1}',p0=distance,p1=unit),mid.x()+8,mid.y()+8,QColor(0,70,80,170))
            if points:
                summary = ui_format('总长: {p0:.2f} {p1} | {p2} 点',p0=data.get('total',0),p1=unit,p2=len(points))
                summary += '\n'+(ui('测量已结束；左键开始下一次。') if state.get('measure_completed') else ui('右键或 Enter 结束本次测量。'))
                draw_textbox(summary,8,max(8,self.height()-70),QColor(0,70,80,220))

        # =============== 全局常驻图层 ===============
        if mode in ["MEASURE_1", "MEASURE_2"] and mouse_pos:
            scr_m = self._to_screen_point(mouse_pos)
            coord_str = get_formatted_coords(mouse_pos)
            draw_textbox(ui_format('📍 当前坐标: {p0}', p0=coord_str), scr_m.x() + 15, scr_m.y() + 15, QColor(0, 188, 212, 200))

        self._paint_mark_layer(painter, state)

        # 全屏十字光标
        if mouse_pos and mode != "IDLE":
            m = self._to_screen_point(mouse_pos)
            if widget_rect.contains(m.toPoint()):
                painter.setPen(QPen(QColor(255, 255, 255, 175), 1))
                if cursor_mode == 1:
                    # 与当前物理坐标轴同向；Qt 屏幕坐标的 Y 轴向下。
                    length = max(widget_rect.width(), widget_rect.height()) * 2.0
                    ux, uy = math.cos(coord_angle), -math.sin(coord_angle)
                    vx, vy = -math.sin(coord_angle), -math.cos(coord_angle)
                    painter.drawLine(
                        QPointF(m.x() - length * ux, m.y() - length * uy),
                        QPointF(m.x() + length * ux, m.y() + length * uy)
                    )
                    painter.drawLine(
                        QPointF(m.x() - length * vx, m.y() - length * vy),
                        QPointF(m.x() + length * vx, m.y() + length * vy)
                    )
                else:
                    painter.drawLine(QPointF(m.x(), 0), QPointF(m.x(), widget_rect.height()))
                    painter.drawLine(QPointF(0, m.y()), QPointF(widget_rect.width(), m.y()))

    def wheelEvent(self, event: QWheelEvent):
        old_coords = self._get_canvas_coords(event.position().toPoint())
        zoom_factor = 1.15 if event.angleDelta().y() > 0 else 1.0 / 1.15
        new_zoom = self.zoom_level * zoom_factor
        if 0.01 <= new_zoom <= 40.0:
            self.zoom_level = new_zoom
            if old_coords:
                cx, cy = old_coords
                lw, lh = self.width(), self.height()
                self.pan_offset_x = event.position().x() - (lw / 2 + (cx - self.canvas_w / 2) * self.zoom_level)
                self.pan_offset_y = event.position().y() - (lh / 2 + (cy - self.canvas_h / 2) * self.zoom_level)
            self.zoom_changed.emit(int(self.zoom_level * 100))
            self.update()
            self.moved_pos.emit(-1.0, -1.0) 

    def mousePressEvent(self, event):
        self.setFocus()
        if event.button() == Qt.MouseButton.RightButton:
            if self.overlay_state.get('mode') in ('MEASURE_1', 'MEASURE_2'):
                self.measurement_finished.emit()
                event.accept()
                return
            self.is_panning = True
            self.last_mouse_pos = event.position()
        elif event.button() == Qt.MouseButton.LeftButton:
            self._last_left_was_click = False
            mode = self.overlay_state.get('mode')
            editing = mode == 'MARK_PAUSED' or (mode == 'MARKING' and bool(event.modifiers() & Qt.KeyboardModifier.ControlModifier))
            if editing:
                index = self._hit_point(event.position())
                self.point_pressed.emit(index)
                self._dragging_point = index >= 0
                self._left_down_pos = event.position()
                self._drag_started = False
            else:
                coords = self._get_canvas_coords(event.position())
                if coords:
                    self._last_left_was_click = True
                    self.clicked_pos.emit(*coords)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.RightButton:
            self.is_panning = False
        elif event.button() == Qt.MouseButton.LeftButton and self._dragging_point:
            coords = self._get_canvas_coords(event.position())
            if coords and self._drag_started:
                self.point_dragged.emit(*coords)
            self._dragging_point = False
            self.point_released.emit()

    def mouseMoveEvent(self, event):
        if self._dragging_point:
            if not self._drag_started and (event.position() - self._left_down_pos).manhattanLength() >= QApplication.startDragDistance():
                self._drag_started = True
            coords = self._get_canvas_coords(event.position())
            if coords and self._drag_started:
                self.point_dragged.emit(*coords)
        elif self.is_panning and self.last_mouse_pos is not None:
            delta = event.position() - self.last_mouse_pos
            self.pan_offset_x += delta.x()
            self.pan_offset_y += delta.y()
            self.last_mouse_pos = event.position()
            self.update()
            self.moved_pos.emit(-1.0, -1.0)
        else:
            coords = self._get_canvas_coords(event.position())
            if coords:
                self.moved_pos.emit(*coords)

    def leaveEvent(self, event):
        self.moved_pos.emit(-1.0, -1.0)
        super().leaveEvent(event)

    def _paint_mark_layer(self, painter, state):
        key = (self.width(), self.height(), self.zoom_level, self.pan_offset_x,
               self.pan_offset_y, state.get('marks_revision'), state.get('selected_point'))
        if self._mark_layer_key != key:
            layer = QPixmap(self.size())
            layer.fill(Qt.GlobalColor.transparent)
            p = QPainter(layer)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            rect = self._view_rect()
            z, ox, oy = self.zoom_level, rect.x(), rect.y()
            def screen(pt):
                return QPointF(ox + pt[0] * z, oy + pt[1] * z)
            for i, trail in enumerate(state.get('all_trails', [])):
                if not trail:
                    continue
                polygon = QPolygonF([screen(pt) for pt in trail])
                color = QColor(TRACK_COLORS_VIEW[i % len(TRACK_COLORS_VIEW)])
                color.setAlpha(130)
                p.setPen(QPen(color, 1.2))
                p.drawPolyline(polygon)
            selected = state.get('selected_point', -1)
            for group, radius, alpha in [('previous_marks', 3.5, 110), ('current_marks', 5.0, 235)]:
                marks = state.get(group, [])
                paths = [QPainterPath() for _ in TRACK_COLORS_VIEW]
                for i, mark in enumerate(marks):
                    v = screen(mark['pixel_pt'])
                    if not self.rect().adjusted(-20, -20, 20, 20).contains(v.toPoint()):
                        continue
                    path = paths[i % len(paths)]
                    path.moveTo(v.x()-radius, v.y()); path.lineTo(v.x()+radius, v.y())
                    path.moveTo(v.x(), v.y()-radius); path.lineTo(v.x(), v.y()+radius)
                for i, path in enumerate(paths):
                    color = QColor(TRACK_COLORS_VIEW[i]); color.setAlpha(alpha)
                    p.setPen(QPen(color, 1.2)); p.drawPath(path)
                p.setFont(QFont('Segoe UI', 9))
                # 密集点集仅显示所选点及最新五点的编号；所有编号、坐标仍完整保存。
                for i, mark in enumerate(marks):
                    if len(marks) > 80 and i != selected and i < len(marks)-5:
                        continue
                    v = screen(mark['pixel_pt'])
                    p.setPen(QColor(mark['color']))
                    p.drawText(v + QPointF(6, -7), ui(mark['text']))
                if group == 'current_marks' and 0 <= selected < len(marks):
                    v = screen(marks[selected]['pixel_pt'])
                    p.setPen(QPen(QColor('#ffffff'), 2)); p.setBrush(Qt.BrushStyle.NoBrush)
                    p.drawEllipse(v, 9, 9)
            p.end()
            self._mark_layer, self._mark_layer_key = layer, key
        painter.drawPixmap(0, 0, self._mark_layer)

    def _hit_point(self, pos):
        coords = self._get_canvas_coords(pos)
        if coords is None:
            return -1
        marks = self.overlay_state.get('current_marks', [])
        if not marks:
            return -1
        xy = np.asarray([d['pixel_pt'] for d in marks], dtype=float)
        dist = np.sum((xy - coords) ** 2, axis=1)
        index = int(np.argmin(dist))
        return index if dist[index] <= (10 / self.zoom_level) ** 2 else -1

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging_point = False
            self.fullscreen_requested.emit(self._last_left_was_click)
            event.accept()
        else:
            super().mouseDoubleClickEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and self.overlay_state.get('mode') in ('MEASURE_1', 'MEASURE_2'):
            self.measurement_finished.emit()
            event.accept()
            return
        super().keyPressEvent(event)


# ==========================================================
# 3. 主界面 (MainWindow)
# ==========================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(ui("Open Tracker V3.9"))
        self.resize(1600, 900)

        # 视频状态参数
        self.video_path = ""
        self.total_frames = 0
        self.fps = 30.0
        self.video_w = 0
        self.video_h = 0
        self.current_frame_idx = -1
        self.base_frame = None       
        self.base_qimage = None
        self.adjusted_qimage = None
        self.is_slider_dragging = False

        # 初始化后台独立解码线程
        self.decoder_worker = VideoDecoderWorker()
        self.decoder_worker.video_info_loaded.connect(self.on_video_info_loaded)
        self.decoder_worker.frame_loaded.connect(self.on_frame_loaded_from_worker)
        self.decoder_worker.error_signal.connect(self.on_worker_error)
        self.decoder_worker.start()

        self.state = "IDLE" 
        self.calib_pt1 = None
        self.calib_line = None
        self.scale_factor = 1.0 
        self.physics_unit = "单位"
        self.grid_size = 10.0   
        self.show_overlays = True 
        
        self.coord_origin = None   
        self.coord_angle = 0.0 
        self.coord_pending_origin = None
        self.coord_return_state = "IDLE"
        
        self.measure_pt1 = None
        self.measure_line = None
        self.measure_points = []
        self.measure_completed = False
        
        self.mouse_curr_pos = None 
        self.brightness = 0
        self.contrast = 1.0
        self.points_per_frame = 1 
        self.cursor_mode = 0  # 0: 屏幕水平/竖直；1: 跟随当前坐标轴

        self.aux_lines = []
        self.aux_current_points = []
        self.aux_intersections = []
        self.aux_midpoints = []

        self.marks = {} 
        self.first_time_sec = None 
        
        # 播放定时器（仅负责从 Buffer 提帧）
        self.play_timer = QTimer()
        self.play_timer.timeout.connect(self.play_next_frame)
        
        self.curves_x = []
        self.curves_y = []
        self.highlights_x = []
        self.highlights_y = []

        # 全局撤回/还原历史。捕捉数据按帧保存差量，避免长视频反复复制全部数据。
        self.undo_stack = []
        self.redo_stack = []
        self._history_restoring = False
        self._adjustment_history_before = None
        self._adjustment_history_timer = QTimer(self)
        self._adjustment_history_timer.setSingleShot(True)
        self._adjustment_history_timer.timeout.connect(self._finish_adjustment_history)

        self.sharpness = 0.0
        self.extra_adjustments = dict(EXTRA_DEFAULTS)
        self.theme = "dark"
        self.language = "zh"
        self._browser_width = 280
        self._browser_resize_start = None
        self._plots_collapsed = False
        self.selected_point = -1
        self._drag_context = None
        self._marks_revision = 0
        self._overlay_cache_key = None
        self._overlay_cache = None
        self._is_fullscreen_video = False
        self._seek_pending = False
        self.auto_advance = True
        self._last_canvas_click = None
        self._plot_timer = QTimer(self)
        self._plot_timer.setSingleShot(True)
        self._plot_timer.timeout.connect(self._flush_plots)
        self._plots_dirty = False
        self._cursor_timer = QTimer(self)
        self._cursor_timer.setSingleShot(True)
        self._cursor_timer.timeout.connect(self.video_cursor_update)
        self.init_ui()
        self.create_file_browser()
        self.create_quality_dialog()
        self.create_edge_handles()
        self.apply_theme("dark")
        self.set_language("zh")
        self._sync_capture_controls()

    def _custom_icon(self, icon_type):
        icon = QIcon()
        colors = ("#344256", "#1267d3", "#a2adba") if self.theme == 'light' else ("#edf2f8", "#ffffff", "#687484")
        for mode, color_hex in zip((QIcon.Mode.Normal, QIcon.Mode.Active, QIcon.Mode.Disabled), colors):
            pix = QPixmap(32, 32)
            pix.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pix)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            color = QColor(color_hex)
            pen = QPen(color, 2.0, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            
            if icon_type == "load":
                painter.drawRoundedRect(4, 12, 24, 14, 2, 2)
                painter.drawLine(4, 12, 12, 12)
                painter.drawLine(12, 12, 16, 6)
                painter.drawLine(16, 6, 28, 6)
                painter.drawLine(28, 6, 28, 12)
            elif icon_type == "calib":
                painter.drawLine(4, 16, 28, 16)
                painter.drawLine(4, 16, 10, 10)
                painter.drawLine(4, 16, 10, 22)
                painter.drawLine(28, 16, 22, 10)
                painter.drawLine(28, 16, 22, 22)
                painter.drawLine(16, 12, 16, 20)
                painter.drawLine(12, 14, 12, 18)
                painter.drawLine(20, 14, 20, 18)
            elif icon_type == "origin":
                painter.drawLine(6, 6, 6, 26)
                painter.drawLine(6, 26, 26, 26)
                painter.drawLine(2, 10, 6, 6)
                painter.drawLine(10, 10, 6, 6)
                painter.drawLine(22, 22, 26, 26)
                painter.drawLine(22, 30, 26, 26)
                painter.setBrush(color)
                painter.drawEllipse(QPointF(6, 26), 2.0, 2.0)
            elif icon_type == "measure":
                painter.translate(16, 16)
                painter.rotate(-45)
                painter.drawRoundedRect(-14, -6, 28, 12, 2, 2)
                painter.drawLine(-8, -6, -8, -2)
                painter.drawLine(-2, -6, -2, -2)
                painter.drawLine(4, -6, 4, -2)
                painter.drawLine(10, -6, 10, -2)
            elif icon_type == "mark":
                painter.drawEllipse(8, 8, 16, 16)
                painter.drawLine(16, 2, 16, 6)
                painter.drawLine(16, 26, 16, 30)
                painter.drawLine(2, 16, 6, 16)
                painter.drawLine(26, 16, 30, 16)
                painter.setBrush(color)
                painter.drawEllipse(QPointF(16, 16), 2.0, 2.0)
            elif icon_type == "aux":
                painter.drawLine(6, 26, 14, 8)
                painter.drawLine(26, 26, 18, 8)
                painter.drawLine(6, 26, 26, 26)
                painter.drawLine(14, 8, 18, 8)
                painter.setPen(QPen(color, 1.5, Qt.PenStyle.DashLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
                painter.drawLine(6, 8, 26, 26)
                painter.drawLine(26, 8, 6, 26)
            elif icon_type == "export":
                painter.drawLine(16, 4, 16, 20)
                painter.drawLine(10, 14, 16, 20)
                painter.drawLine(22, 14, 16, 20)
                painter.drawLine(6, 24, 26, 24)
                painter.drawLine(6, 28, 26, 28)
            elif icon_type == "import":
                painter.drawLine(16, 20, 16, 4)
                painter.drawLine(10, 10, 16, 4)
                painter.drawLine(22, 10, 16, 4)
                painter.drawLine(6, 24, 26, 24)
                painter.drawLine(6, 28, 26, 28)
            elif icon_type == "frame":
                painter.drawRoundedRect(4, 6, 24, 20, 2, 2)
                painter.drawEllipse(8, 10, 3, 3)
                painter.setPen(QPen(color, 1.5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
                painter.drawLine(4, 20, 12, 14)
                painter.drawLine(12, 14, 18, 20)
                painter.drawLine(16, 18, 20, 14)
                painter.drawLine(20, 14, 28, 22)
            elif icon_type == "reset":
                painter.drawRoundedRect(4, 4, 24, 24, 2, 2)
                painter.drawLine(16, 16, 8, 8)
                painter.drawLine(8, 8, 14, 8)
                painter.drawLine(8, 8, 8, 14)
                painter.drawLine(16, 16, 24, 24)
                painter.drawLine(24, 24, 18, 24)
                painter.drawLine(24, 24, 24, 18)
            elif icon_type == "trail":
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawEllipse(QPointF(7, 16), 2.0, 2.0)
                painter.drawEllipse(QPointF(16, 16), 3.0, 3.0)
                painter.drawEllipse(QPointF(25, 16), 4.0, 4.0)
                painter.setPen(QPen(color, 1.5, Qt.PenStyle.DashLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
                painter.drawLine(7, 16, 25, 16)
            elif icon_type == "cursor":
                painter.drawLine(16, 3, 16, 29)
                painter.drawLine(3, 16, 29, 16)
                painter.drawEllipse(QPointF(16, 16), 4.0, 4.0)
            elif icon_type == "undo":
                painter.drawArc(QRectF(7, 8, 20, 18), 25 * 16, 245 * 16)
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPolygon(QPolygonF([
                    QPointF(5, 13), QPointF(13, 8), QPointF(12, 17)
                ]))
            elif icon_type == "redo":
                painter.drawArc(QRectF(5, 8, 20, 18), -90 * 16, 245 * 16)
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPolygon(QPolygonF([
                    QPointF(27, 13), QPointF(19, 8), QPointF(20, 17)
                ]))
            elif icon_type == "play":
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPolygon(QPolygonF([QPointF(10, 6), QPointF(24, 16), QPointF(10, 26)]))
            elif icon_type == "pause":
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawRect(8, 8, 5, 16)
                painter.drawRect(19, 8, 5, 16)
            elif icon_type == "prev":
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPolygon(QPolygonF([QPointF(20, 8), QPointF(10, 16), QPointF(20, 24)]))
                painter.drawRect(6, 8, 3, 16)
            elif icon_type == "next":
                painter.setBrush(color)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPolygon(QPolygonF([QPointF(12, 8), QPointF(22, 16), QPointF(12, 24)]))
                painter.drawRect(23, 8, 3, 16)
            painter.end()
            icon.addPixmap(pix, mode)
        return icon

    def create_actions(self):
        self.act_load = QAction(self._custom_icon("load"), ui("加载视频"), self)
        self.act_load.setShortcut(QKeySequence.StandardKey.Open)
        self.act_load.triggered.connect(self.load_video)

        self.act_undo = QAction(self._custom_icon("undo"), ui("撤回"), self)
        self.act_undo.setShortcuts(QKeySequence.keyBindings(QKeySequence.StandardKey.Undo))
        self.act_undo.setStatusTip(ui("撤回上一步操作 (Ctrl+Z)"))
        self.act_undo.triggered.connect(self.undo_last_action)
        self.act_undo.setEnabled(False)

        self.act_redo = QAction(self._custom_icon("redo"), ui("还原"), self)
        self.act_redo.setShortcuts(QKeySequence.keyBindings(QKeySequence.StandardKey.Redo))
        self.act_redo.setStatusTip(ui("还原下一步操作 (Ctrl+Y / Ctrl+Shift+Z)"))
        self.act_redo.triggered.connect(self.redo_last_action)
        self.act_redo.setEnabled(False)

        self.act_export = QAction(self._custom_icon("export"), ui("导出 Excel"), self)
        self.act_export.setShortcut(QKeySequence.StandardKey.SaveAs)
        self.act_export.triggered.connect(self.export_to_excel)

        self.act_export_frame = QAction(self._custom_icon("frame"), ui("导出当前帧"), self)
        self.act_export_frame.triggered.connect(self.export_current_frame)

        self.act_import = QAction(self._custom_icon("import"), ui("导入 Excel"), self)
        self.act_import.setShortcut(QKeySequence("Ctrl+Shift+O"))
        self.act_import.triggered.connect(self.import_from_excel)

        self.act_calib = QAction(self._custom_icon("calib"), ui("标定尺寸"), self)
        self.act_calib.triggered.connect(lambda: self.start_state_with_history("CALIBRATING_1"))

        self.act_coord = QAction(self._custom_icon("origin"), ui("设置坐标"), self)
        self.act_coord.triggered.connect(self.start_coordinate_setting)
        
        self.act_measure = QAction(self._custom_icon("measure"), ui("测量查询"), self)
        self.act_measure.triggered.connect(self.start_measurement)

        self.act_mark = QAction(self._custom_icon("mark"), ui("开始捕捉"), self)
        self.act_mark.triggered.connect(self.start_marking_mode)
        
        self.act_aux_line = QAction(self._custom_icon("aux"), ui("透视辅助线"), self)
        self.act_aux_line.triggered.connect(lambda: self.start_state_with_history("DRAW_AUX"))

        self.act_reset_view = QAction(self._custom_icon("reset"), ui("画面重置"), self)
        self.act_reset_view.triggered.connect(self.reset_label_view)

        self.act_escape = QAction(ui('退出视频全屏'), self)
        self.act_escape.setShortcut(QKeySequence('Escape'))
        self.act_escape.triggered.connect(self.exit_video_fullscreen)
        self.addAction(self.act_escape)
        self.act_play_shortcut = QAction(self)
        self.act_play_shortcut.setShortcut(QKeySequence('Space'))
        self.act_play_shortcut.triggered.connect(lambda: self.toggle_play() if self.total_frames else None)
        self.addAction(self.act_play_shortcut)
        self.act_pause_capture = QAction(ui('暂停捕捉'), self)
        self.act_pause_capture.setShortcut(QKeySequence('P'))
        self.act_pause_capture.triggered.connect(self.toggle_capture_pause)
        self.act_pause_capture.setEnabled(False)
        for action in (self.act_pause_capture, self.act_undo, self.act_redo):
            self.addAction(action)  # 工具栏/菜单隐藏后，窗口级快捷键仍有效。
        self.act_quality = QAction(ui('画面调节'), self)
        self.act_quality.triggered.connect(self.show_quality_dialog)
        self.act_quality.setEnabled(False)
        self.act_theme = QAction(ui('浅色界面'), self)
        self.act_theme.triggered.connect(self.toggle_theme)
        self.act_language = QAction(ui('English'), self)
        self.act_language.triggered.connect(self.toggle_language)
        self.act_exit = QAction(ui("退出"), self)
        self.act_exit.setShortcut(QKeySequence.StandardKey.Quit)
        self.act_exit.triggered.connect(self.close)

        self.video_actions = [
            self.act_calib, self.act_coord, self.act_measure, self.act_mark,
            self.act_aux_line, self.act_export, self.act_import,
            self.act_export_frame, self.act_reset_view, self.act_quality,
        ]
        for action in self.video_actions:
            action.setEnabled(False)

    def create_menus(self):
        file_menu = self.menuBar().addMenu(ui("文件"))
        file_menu.addAction(self.act_load)
        file_menu.addAction(self.act_export)
        file_menu.addAction(self.act_import)
        file_menu.addAction(self.act_export_frame)
        file_menu.addSeparator()
        file_menu.addAction(self.act_exit)

        edit_menu = self.menuBar().addMenu(ui("编辑"))
        edit_menu.addAction(self.act_undo)
        edit_menu.addAction(self.act_redo)

        track_menu = self.menuBar().addMenu(ui("分析"))
        track_menu.addAction(self.act_calib)
        track_menu.addAction(self.act_coord)
        track_menu.addAction(self.act_measure)
        track_menu.addAction(self.act_mark)
        
        tool_menu = self.menuBar().addMenu(ui("工具"))
        tool_menu.addAction(self.act_aux_line)

        view_menu = self.menuBar().addMenu(ui("视图"))
        self.view_menu = view_menu
        view_menu.addAction(self.act_reset_view)
        view_menu.addAction(self.act_theme)
        view_menu.addAction(self.act_language)
        tool_menu.addAction(self.act_quality)

    def create_toolbars(self):
        toolbar = QToolBar(ui("主工具栏"), self)
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(20, 20))
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        
        toolbar.addAction(self.act_load)
        toolbar.addSeparator()

        toolbar.addAction(self.act_undo)
        toolbar.addAction(self.act_redo)
        toolbar.addSeparator()
        
        toolbar.addAction(self.act_calib)
        toolbar.addAction(self.act_coord)
        toolbar.addAction(self.act_aux_line)

        self.btn_cursor = QToolButton(self)
        self.btn_cursor.setIcon(self._custom_icon("cursor"))
        self.btn_cursor.setText(ui("光标模式"))
        self.btn_cursor.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.btn_cursor.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_cursor.setEnabled(False)

        cursor_menu = QMenu(self)
        self.cursor_actions = QActionGroup(self)
        self.act_cursor_screen = QAction(ui("屏幕水平/竖直（默认）"), self, checkable=True)
        self.act_cursor_axes = QAction(ui("跟随当前坐标轴"), self, checkable=True)
        self.act_cursor_screen.setChecked(True)
        self.cursor_actions.addAction(self.act_cursor_screen)
        self.cursor_actions.addAction(self.act_cursor_axes)
        cursor_menu.addAction(self.act_cursor_screen)
        cursor_menu.addAction(self.act_cursor_axes)
        self.btn_cursor.setMenu(cursor_menu)
        self.cursor_actions.triggered.connect(self.on_cursor_mode_changed)
        toolbar.addWidget(self.btn_cursor)
        toolbar.addSeparator()
        
        toolbar.addAction(self.act_measure)
        toolbar.addSeparator()
        
        toolbar.addAction(self.act_mark)
        self.act_pause_capture.setIcon(self._custom_icon("pause"))
        toolbar.addAction(self.act_pause_capture)
        
        self.btn_trail = QToolButton(self)
        self.btn_trail.setIcon(self._custom_icon("trail"))
        self.btn_trail.setText(ui("轨迹模式"))
        self.btn_trail.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.btn_trail.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_trail.setEnabled(False)

        trail_menu = QMenu(self)
        self.trail_actions = QActionGroup(self)
        self.trail_mode = 1 
        self.act_trail_none = QAction(ui("无轨迹"), self, checkable=True)
        self.act_trail_local = QAction(ui("局部轨迹 (近10点)"), self, checkable=True)
        self.act_trail_all = QAction(ui("全部轨迹"), self, checkable=True)
        self.act_trail_local.setChecked(True)
        self.trail_actions.addAction(self.act_trail_none)
        self.trail_actions.addAction(self.act_trail_local)
        self.trail_actions.addAction(self.act_trail_all)
        trail_menu.addAction(self.act_trail_none)
        trail_menu.addAction(self.act_trail_local)
        trail_menu.addAction(self.act_trail_all)
        self.btn_trail.setMenu(trail_menu)

        self.trail_actions.triggered.connect(self.on_trail_changed)
        toolbar.addWidget(self.btn_trail)
        toolbar.addSeparator()
        
        toolbar.addAction(self.act_import)
        toolbar.addAction(self.act_export)
        toolbar.addAction(self.act_export_frame)
        toolbar.addSeparator()
        
        toolbar.addAction(self.act_reset_view)
        toolbar.addAction(self.act_theme)
        toolbar.addAction(self.act_language)
        self.main_toolbar = toolbar
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, toolbar)

    # ==========================================================
    # 全局撤回/还原
    # ==========================================================

    def _capture_history_state(self):
        """捕获除大体量 marks 外的完整可编辑状态。"""
        return {
            'state': self.state,
            'calib_pt1': copy.deepcopy(self.calib_pt1),
            'calib_line': copy.deepcopy(self.calib_line),
            'scale_factor': self.scale_factor,
            'physics_unit': self.physics_unit,
            'grid_size': self.grid_size,
            'show_overlays': self.show_overlays,
            'coord_origin': copy.deepcopy(self.coord_origin),
            'coord_angle': self.coord_angle,
            'coord_pending_origin': copy.deepcopy(self.coord_pending_origin),
            'coord_return_state': self.coord_return_state,
            'measure_pt1': copy.deepcopy(self.measure_pt1),
            'measure_line': copy.deepcopy(self.measure_line),
            'measure_points': copy.deepcopy(self.measure_points),
            'measure_completed': self.measure_completed,
            'aux_lines': copy.deepcopy(self.aux_lines),
            'aux_current_points': copy.deepcopy(self.aux_current_points),
            'aux_intersections': copy.deepcopy(self.aux_intersections),
            'aux_midpoints': copy.deepcopy(self.aux_midpoints),
            'first_time_sec': self.first_time_sec,
            'points_per_frame': self.points_per_frame,
            'trail_mode': self.trail_mode,
            'cursor_mode': self.cursor_mode,
            'brightness': self.brightness,
            'contrast': self.contrast,
            'sharpness': self.sharpness,
            'extra_adjustments': dict(self.extra_adjustments),
            'selected_point': self.selected_point,
            'zoom_level': self.video_widget.zoom_level,
            'pan_offset_x': self.video_widget.pan_offset_x,
            'pan_offset_y': self.video_widget.pan_offset_y,
        }

    def _commit_history(self, before_state, description, mark_change=None):
        if self._history_restoring or before_state is None:
            return
        after_state = self._capture_history_state()
        if before_state == after_state and mark_change is None:
            return
        self.undo_stack.append({
            'before': before_state,
            'after': after_state,
            'mark_change': mark_change,
            'description': description,
        })
        self.redo_stack.clear()
        self._update_history_actions()

    def _update_history_actions(self):
        if hasattr(self, 'act_undo'):
            self.act_undo.setEnabled(bool(self.undo_stack))
            self.act_redo.setEnabled(bool(self.redo_stack))

    def _apply_mark_change(self, change, use_after):
        if not change:
            return
        side = 'after' if use_after else 'before'
        if change['kind'] == 'point':
            frame_idx, index = change['frame_idx'], change['index']
            value = change[side]
            points = self.marks.setdefault(frame_idx, [])
            if change['operation'] == 'append':
                if use_after:
                    points.insert(index, copy.deepcopy(value))
                elif index < len(points):
                    points.pop(index)
                if not points and not change.get('before_exists', False):
                    self.marks.pop(frame_idx, None)
            else:
                points[index] = copy.deepcopy(value)
        elif change['kind'] == 'frame':
            frame_idx = change['frame_idx']
            exists = change[f'{side}_exists']
            if exists:
                self.marks[frame_idx] = copy.deepcopy(change[side])
            else:
                self.marks.pop(frame_idx, None)
        elif change['kind'] == 'all':
            self.marks = copy.deepcopy(change[side])

    def _apply_history_state(self, snapshot):
        self.state = snapshot['state']
        self.calib_pt1 = copy.deepcopy(snapshot['calib_pt1'])
        self.calib_line = copy.deepcopy(snapshot['calib_line'])
        self.scale_factor = snapshot['scale_factor']
        self.physics_unit = snapshot['physics_unit']
        self.grid_size = snapshot['grid_size']
        self.show_overlays = snapshot['show_overlays']
        self.coord_origin = copy.deepcopy(snapshot['coord_origin'])
        self.coord_angle = snapshot['coord_angle']
        self.coord_pending_origin = copy.deepcopy(snapshot['coord_pending_origin'])
        self.coord_return_state = snapshot['coord_return_state']
        self.measure_pt1 = copy.deepcopy(snapshot['measure_pt1'])
        self.measure_line = copy.deepcopy(snapshot['measure_line'])
        self.measure_points = copy.deepcopy(snapshot.get('measure_points', []))
        self.measure_completed = snapshot.get('measure_completed', False)
        self.aux_lines = copy.deepcopy(snapshot['aux_lines'])
        self.aux_current_points = copy.deepcopy(snapshot['aux_current_points'])
        self.aux_intersections = copy.deepcopy(snapshot['aux_intersections'])
        self.aux_midpoints = copy.deepcopy(snapshot['aux_midpoints'])
        self.first_time_sec = snapshot['first_time_sec']
        self.points_per_frame = snapshot['points_per_frame']
        self.trail_mode = snapshot['trail_mode']
        self.cursor_mode = snapshot['cursor_mode']
        self.brightness = snapshot['brightness']
        self.contrast = snapshot['contrast']
        self.sharpness = snapshot.get('sharpness', 0.0)
        self.extra_adjustments = dict(EXTRA_DEFAULTS, **snapshot.get('extra_adjustments', {}))
        self.selected_point = snapshot.get('selected_point', -1)

        self.video_widget.zoom_level = snapshot['zoom_level']
        self.video_widget.pan_offset_x = snapshot['pan_offset_x']
        self.video_widget.pan_offset_y = snapshot['pan_offset_y']

        self.chk_show_overlays.blockSignals(True)
        self.chk_show_overlays.setChecked(self.show_overlays)
        self.chk_show_overlays.blockSignals(False)
        self._set_adjustment_controls()

        trail_action = {
            0: self.act_trail_none,
            1: self.act_trail_local,
            2: self.act_trail_all,
        }.get(self.trail_mode, self.act_trail_local)
        trail_action.setChecked(True)

        cursor_action = (
            self.act_cursor_axes if self.cursor_mode == 1 else self.act_cursor_screen
        )
        cursor_action.setChecked(True)

        self._send_adjustments()

        self._sync_quality_values()
        # 固定数量图元；脏数据刷新也会清空“撤销到无数据”后的旧曲线。
        self._init_plot_curves(self.points_per_frame)
        self.update_chart_labels()
        self.update_plots()
        self._refresh_state_status()
        self.render_frame()

    def _apply_history_entry(self, entry, use_after):
        if self.play_timer.isActive():
            self.toggle_play()
        self._history_restoring = True
        try:
            self._apply_mark_change(entry.get('mark_change'), use_after)
            snapshot = entry['after'] if use_after else entry['before']
            self._apply_history_state(snapshot)
            change = entry.get('mark_change')
            if change and 'frame_idx' in change:
                self.seek_frame(change['frame_idx'], absolute=True)
        finally:
            self._history_restoring = False

    def _rollback_uncommitted_action(self, before_state, before_marks):
        """导入失败或用户取消时，原子恢复操作开始前的状态。"""
        self._history_restoring = True
        try:
            self.marks = copy.deepcopy(before_marks)
            self._apply_history_state(before_state)
        finally:
            self._history_restoring = False

    def undo_last_action(self):
        self.finish_point_drag()
        self._finish_adjustment_history()
        if not self.undo_stack:
            return
        entry = self.undo_stack.pop()
        self.redo_stack.append(entry)
        self._apply_history_entry(entry, use_after=False)
        self._update_history_actions()

    def redo_last_action(self):
        self.finish_point_drag()
        self._finish_adjustment_history()
        if not self.redo_stack:
            return
        entry = self.redo_stack.pop()
        self.undo_stack.append(entry)
        self._apply_history_entry(entry, use_after=True)
        self._update_history_actions()

    def start_state_with_history(self, target_state):
        if target_state == "MEASURE_1":
            self.start_measurement()
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        self.set_state(target_state)
        self._commit_history(before, ui_format('进入 {p0}', p0=target_state))

    def on_trail_changed(self, action):
        if self._history_restoring:
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        if action == self.act_trail_none:
            self.trail_mode = 0
        elif action == self.act_trail_local:
            self.trail_mode = 1
        elif action == self.act_trail_all:
            self.trail_mode = 2
        self.render_frame()
        self._commit_history(before, '切换轨迹模式')

    def on_cursor_mode_changed(self, action):
        if self._history_restoring:
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        self.cursor_mode = 1 if action == self.act_cursor_axes else 0
        self.render_frame()
        self._commit_history(before, '切换光标模式')

    def start_marking_mode(self):
        pts, ok = QInputDialog.getInt(self, ui('开启捕捉模式'), ui('每帧目标点数（按轮廓顺序点击，支持数百点）：'), self.points_per_frame, 1, 4000, 1)
        if not ok:
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        if self.play_timer.isActive():
            self.toggle_play()
        self.points_per_frame = pts  # 修改配额不会删除任何既有点。
        self.selected_point = -1
        self.set_state('MARKING')
        self._commit_history(before, '开始捕捉/修改目标点数')

    def start_coordinate_setting(self):
        """允许在连续捕捉过程中临时进入坐标系重定义，并在结束后恢复捕捉。"""
        self._finish_adjustment_history()
        before = self._capture_history_state()
        if self.play_timer.isActive():
            self.toggle_play()
        self.coord_return_state = self.state if self.state in ("MARKING", "MARK_PAUSED") else "IDLE"
        self.coord_pending_origin = None
        self.set_state("SET_COORD_1")
        self._commit_history(before, '开始设置坐标系')

    def _init_plot_curves(self, num_points):
        if self.curves_x:
            return
        for plot, curves, colors, highlights in [
                (self.plot_x, self.curves_x, TRACK_COLORS_X, self.highlights_x),
                (self.plot_y, self.curves_y, TRACK_COLORS_Y, self.highlights_y)]:
            for color in colors:
                curve = plot.plot(pen=pg.mkPen(color, width=1.2), symbol='o', symbolSize=4,
                                  symbolBrush=color, symbolPen=None, connect='finite')
                curve.setClipToView(True)
                curve.sigPointsClicked.connect(self.on_plot_point_clicked)
                curves.append(curve)
            hl = plot.plot([], [], pen=None, symbol='star', symbolSize=11,
                           symbolBrush='#ffffff', symbolPen=pg.mkPen('#ffffff'))
            hl.setZValue(10)
            highlights.append(hl)

    def init_ui(self):
        self.create_actions()
        self.create_menus()
        self.create_toolbars()

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        root_layout = QVBoxLayout(main_widget)
        self.root_layout = root_layout
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(10)

        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)
        left_layout = QVBoxLayout()
        left_layout.setSpacing(0)
        self.video_widget = VideoWidget() 
        self.video_widget.clicked_pos.connect(self.on_video_clicked)
        self.video_widget.moved_pos.connect(self.on_video_moved)
        self.video_widget.point_pressed.connect(self.select_capture_point)
        self.video_widget.point_dragged.connect(self.preview_point_drag)
        self.video_widget.point_released.connect(self.finish_point_drag)
        self.video_widget.fullscreen_requested.connect(self.on_video_double_clicked)
        self.video_widget.measurement_finished.connect(self.finish_measurement)
        left_layout.addWidget(self.video_widget, stretch=1) 

        self.plot_panel = QWidget()
        right_layout = QVBoxLayout(self.plot_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)
        
        self.plot_x = pg.PlotWidget(title=ui("X轴位移 (依自定义坐标系)"))
        self.plot_x.showGrid(x=True, y=True, alpha=0.3)

        self.plot_y = pg.PlotWidget(title=ui("Y轴位移 (依自定义坐标系)"))
        self.plot_y.showGrid(x=True, y=True, alpha=0.3)
        
        self._init_plot_curves(self.points_per_frame)
        self.update_chart_labels()
        
        right_layout.addWidget(self.plot_x)
        right_layout.addWidget(self.plot_y)

        content_layout.addLayout(left_layout, stretch=2)
        content_layout.addWidget(self.plot_panel, stretch=1)
        root_layout.addLayout(content_layout, stretch=1)

        bottom_frame = QFrame()
        self.playback_bar = bottom_frame
        bottom_frame.setObjectName("controlBar")
        bottom_layout = QHBoxLayout(bottom_frame)
        bottom_layout.setContentsMargins(10, 8, 10, 8)
        bottom_layout.setSpacing(8)

        self.btn_play = QPushButton(ui("播放"))
        self.btn_play.setIcon(self._custom_icon("play"))
        self.btn_prev = QPushButton(ui("上一帧"))
        self.btn_prev.setIcon(self._custom_icon("prev"))
        self.btn_next = QPushButton(ui("下一帧"))
        self.btn_next.setIcon(self._custom_icon("next"))

        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setEnabled(False)
        self.slider.setTracking(False) 

        self.label_frame_text = QLabel(ui("帧:"))
        self.spin_frame = QSpinBox()
        self.spin_frame.setRange(0, 0)
        self.spin_frame.setEnabled(False)
        self.spin_frame.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.spin_frame.setMinimumWidth(92)
        self.label_frame_total = QLabel(ui("/ 0"))
        self.label_step = QLabel(ui("步长:"))
        self.spin_step = QSpinBox()
        self.spin_step.setRange(1, 100)
        self.spin_step.setValue(1)
        self.spin_step.setSuffix(ui(" 帧"))
        self.spin_step.setEnabled(False)

        self.btn_quality = QPushButton(ui('画面调节'))
        self.btn_quality.clicked.connect(self.show_quality_dialog)
        self.btn_quality.setEnabled(False)
        self.chk_show_overlays = QCheckBox(ui("显示辅助控件"))
        self.chk_show_overlays.setChecked(True)
        self.chk_show_overlays.setEnabled(False)

        self.btn_play.clicked.connect(self.toggle_play)
        self.btn_prev.clicked.connect(lambda: self.seek_frame(-self.spin_step.value()))
        self.btn_next.clicked.connect(lambda: self.seek_frame(self.spin_step.value()))
        self.slider.sliderPressed.connect(self.on_slider_pressed)
        self.slider.valueChanged.connect(self.on_slider_value_changed)
        self.slider.sliderReleased.connect(self.on_slider_released)
        self.spin_frame.editingFinished.connect(self.on_frame_spin_finished)
        self.chk_show_overlays.toggled.connect(self.on_show_overlays_toggled)

        for widget in [self.btn_play, self.btn_prev, self.btn_next]: widget.setEnabled(False)

        bottom_layout.addWidget(self.btn_play)
        bottom_layout.addWidget(self.btn_prev)
        bottom_layout.addWidget(self.btn_next)
        bottom_layout.addWidget(self.slider, stretch=1)
        bottom_layout.addWidget(self.label_frame_text)
        bottom_layout.addWidget(self.spin_frame)
        bottom_layout.addWidget(self.label_frame_total)
        bottom_layout.addSpacing(10)
        bottom_layout.addWidget(self.label_step)
        bottom_layout.addWidget(self.spin_step)
        bottom_layout.addWidget(self.btn_quality)
        bottom_layout.addWidget(self.chk_show_overlays)
        root_layout.addWidget(bottom_frame)
        status_frame = QFrame()
        self.status_panel = status_frame
        status_frame.setObjectName("statusBar")
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(12, 8, 12, 8)
        self.status_label = QLabel(ui("状态: 等待加载视频..."))
        self.status_label.setMinimumWidth(0)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.status_label.setObjectName("statusLabel")
        status_layout.addWidget(self.status_label)
        root_layout.addWidget(status_frame)

    # ==========================================================
    # 异步多线程交互与解码回调
    # ==========================================================

    def load_video(self):
        filename, _ = QFileDialog.getOpenFileName(self, ui('选择视频文件'), ui(self.browser_root),
                                                 'Video Files (' + ' '.join('*'+s for s in VIDEO_EXTENSIONS) + ');;All Files (*)')
        if filename:
            self.open_video_path(filename)

    def on_video_info_loaded(self, info):
        if info['generation'] != self.decoder_worker.generation:
            return
        self.video_path = info['path']
        self.selected_point = -1
        self._drag_context = None
        self._overlay_cache_key = None
        self._marks_revision += 1
        # 新视频建立独立历史，撤回边界止于该视频刚打开的初始状态。
        self._adjustment_history_timer.stop()
        self._adjustment_history_before = None
        self.undo_stack.clear()
        self.redo_stack.clear()
        self._update_history_actions()

        self.total_frames = info["total_frames"]
        self.fps = info["fps"]
        self.video_w = info["width"]
        self.video_h = info["height"]
        
        self.video_widget.canvas_w = self.video_w
        self.video_widget.canvas_h = self.video_h
        
        self.marks = {}
        self.scale_factor = 1.0
        self.coord_origin = None
        self.coord_angle = 0.0
        self.coord_pending_origin = None
        self.coord_return_state = "IDLE"
        self.calib_pt1 = None
        self.calib_line = None
        self.measure_pt1 = None
        self.measure_line = None
        self.measure_points = []
        self.measure_completed = False
        self.clear_aux_lines()
        self.physics_unit = "单位"
        self.grid_size = 10.0
        self.show_overlays = True
        self.chk_show_overlays.blockSignals(True)
        self.chk_show_overlays.setChecked(True)
        self.chk_show_overlays.blockSignals(False)
        self.trail_mode = 1
        self.act_trail_local.setChecked(True)
        self.cursor_mode = 0
        self.act_cursor_screen.setChecked(True)
        self.current_frame_idx = -1 
        self.base_frame = None
        self.base_qimage = None
        self.adjusted_qimage = None
        self.first_time_sec = None
        
        self.points_per_frame = 1 
        self._init_plot_curves(self.points_per_frame) 
        self.update_chart_labels()
        self.update_plots()
        
        self.slider.blockSignals(True)
        self.slider.setRange(0, self.total_frames - 1)
        self.slider.setValue(0)
        self.slider.blockSignals(False)
        self.spin_frame.blockSignals(True)
        self.spin_frame.setRange(0, self.total_frames - 1)
        self.spin_frame.setValue(0)
        self.spin_frame.blockSignals(False)
        self.label_frame_total.setText(ui(f"/ {self.total_frames - 1}"))
        
        for widget in [self.slider, self.btn_play, self.btn_prev, self.btn_next, self.spin_step, self.btn_quality,
                       self.spin_frame, self.slider_brightness, self.slider_contrast, self.slider_sharpness, self.chk_show_overlays,
                       self.btn_trail, self.btn_cursor]:
            widget.setEnabled(True)
        for action in self.video_actions: action.setEnabled(True)
        
        self.reset_adjustments(update_frame=False)
        self.set_state("IDLE")
        self.video_widget.fit_to_view()

    def on_worker_error(self, err_msg):
        self._seek_pending = False
        QMessageBox.critical(self, ui("视频打开错误"), ui(err_msg))
        self.status_label.setText(ui("状态: 加载视频出错"))

    def seek_frame(self, val, absolute=False):
        if not self.total_frames:
            return
        self.finish_point_drag()
        if self.play_timer.isActive():
            self.toggle_play()
        target = val if absolute else self.current_frame_idx + val
        target = max(0, min(int(target), self.total_frames-1))
        self._seek_pending = True
        self.decoder_worker.request_seek(target)

    def on_frame_loaded_from_worker(self, frame_idx, qimg, base_frame, generation):
        if generation != self.decoder_worker.generation:
            return
        changed_frame = frame_idx != self.current_frame_idx
        self.current_frame_idx = frame_idx
        self._seek_pending = False
        if changed_frame:
            self._drag_context = None
        self.base_frame = base_frame
        self.base_qimage = qimg
        self.adjusted_qimage = qimg
        self.video_widget.set_frame(qimg)
        self._finish_seek(frame_idx)
        self.update_highlight()

    def _finish_seek(self, target_idx):
        for control in (self.slider, self.spin_frame):
            control.blockSignals(True)
            control.setValue(target_idx)
            control.blockSignals(False)
        self._refresh_state_status()
        self.render_frame()

    # ==========================================================
    # 播放器控制与定时器
    # ==========================================================

    def toggle_play(self):
        self.finish_point_drag()
        if not self.total_frames:
            return
        if self.play_timer.isActive():
            self.play_timer.stop()
            self.decoder_worker.stop_playback()
            self.btn_play.setIcon(self._custom_icon('play'))
            self.btn_play.setText(ui('播放'))
        else:
            if self.state == 'MARKING':
                self.set_state('MARK_PAUSED')
            step = self.spin_step.value()
            start = self.current_frame_idx + step
            if start >= self.total_frames:
                start = 0
            self._seek_pending = False
            self.decoder_worker.start_playback(start, step)
            self.play_timer.start(self._playback_interval_ms())
            self.btn_play.setIcon(self._custom_icon('pause'))
            self.btn_play.setText(ui('暂停'))

    def _playback_interval_ms(self):
        step_size = max(1, self.spin_step.value())
        return max(16, int(1000 * step_size / max(self.fps, 1.0)))

    def play_next_frame(self):
        if not self.play_timer.isActive():
            return
        try:
            item = self.decoder_worker.buffer_queue.get_nowait()
        except Empty:
            item = None
        if item is not None:
            self.on_frame_loaded_from_worker(*item)
        if (self.current_frame_idx >= self.total_frames-1 or
                (not self.decoder_worker.is_playing and self.decoder_worker.buffer_queue.empty())):
            self.toggle_play()

    def show_quality_dialog(self):
        if not self.total_frames:
            return
        self.quality_dialog.show()
        self.quality_dialog.raise_()
        self.quality_dialog.activateWindow()

    def on_adjustment_changed(self):
        if self._history_restoring:
            return
        if self._adjustment_history_before is None:
            self._adjustment_history_before = self._capture_history_state()
        self.brightness = self.slider_brightness.value()
        self.contrast = self.slider_contrast.value()/100.0
        self.sharpness = self.slider_sharpness.value()/100.0
        self.extra_adjustments = {name: getattr(self, 'slider_'+name).value()/divisor
                                  for name, _, _, _, _, divisor, _ in EXTRA_ADJUSTMENT_SPECS}
        self._send_adjustments()
        self._adjustment_history_timer.start(300)

    def _finish_adjustment_history(self):
        if self._adjustment_history_timer.isActive():
            self._adjustment_history_timer.stop()
        if self._adjustment_history_before is None:
            return
        before = self._adjustment_history_before
        self._adjustment_history_before = None
        self._commit_history(before, '调整画面参数')

    def reset_label_view(self):
        self._finish_adjustment_history()
        before = self._capture_history_state()
        self.video_widget.fit_to_view()
        self.reset_adjustments(update_frame=True)
        self.render_frame()
        self._commit_history(before, '重置画面')

    def reset_adjustments(self, update_frame=True):
        self.brightness, self.contrast, self.sharpness = 0, 1.0, 0.0
        self.extra_adjustments = dict(EXTRA_DEFAULTS)
        self._set_adjustment_controls()
        self._send_adjustments(update_frame)

    def on_slider_pressed(self): self.is_slider_dragging = True

    def on_slider_value_changed(self, value):
        self.spin_frame.blockSignals(True)
        self.spin_frame.setValue(value)
        self.spin_frame.blockSignals(False)
        if not self.is_slider_dragging: self.seek_frame(value, absolute=True)

    def on_slider_released(self):
        self.is_slider_dragging = False
        self.seek_frame(self.slider.value(), absolute=True)

    def on_frame_spin_finished(self):
        val = self.spin_frame.value()
        self.slider.blockSignals(True)
        self.slider.setValue(val)
        self.slider.blockSignals(False)
        self.seek_frame(val, absolute=True)

    def on_show_overlays_toggled(self, checked):
        if self._history_restoring:
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        self.show_overlays = checked
        self.render_frame()
        self._commit_history(before, '切换辅助控件')

    def on_plot_point_clicked(self, item, points, *args):
        if not points:
            return
        data = points[0].data()
        if not isinstance(data, dict):
            return
        if self.play_timer.isActive():
            self.toggle_play()
        self.selected_point = data['point']
        if self.state == 'MARKING':
            self.set_state('MARK_PAUSED')
        self.seek_frame(data['frame'], absolute=True)

    def _refresh_state_status(self):
        pts = self.points_per_frame
        curr_len = len(self.marks.get(self.current_frame_idx, [])) if self.current_frame_idx in self.marks else 0
        
        msgs = {
            "IDLE": ui("空闲。滚轮缩放，右键按住拖动画面，左下角可点击播放/暂停。"),
            "CALIBRATING_1": ui("【尺寸标定】: 请点击定标尺的【起点】 (自动吸附辅助交点/中点)"),
            "CALIBRATING_2": ui("【尺寸标定】: 请点击定标尺的【终点】 (自动吸附辅助交点/中点)"),
            "SET_COORD_1": ui("【建立坐标系】: 请点击设定【坐标原点 (0,0)】 (自动吸附辅助点)"),
            "SET_COORD_2": ui("【建立坐标系】: 请点击画面确定【X轴正方向】(随后会弹窗供你微调角度)"),
            "MEASURE_1": ui("【测量/查询】: 连续测量模式。请点击起点测量距离 (标定后显示真实物理尺寸)。"),
            "MEASURE_2": ui("【测量/查询】: 请点击终点以固定测量线段，随后可继续下一次测量。"),
            "MARKING": ui_format('【捕捉模式】: 设置为每帧 {p0} 个点。请点击目标（当前第 {p1}/{p2} 个点）', p0=pts, p1=min(curr_len+1, pts), p2=pts),
            "MARK_PAUSED": ui("【捕捉已暂停】左键选中并拖动当前帧的点；P 继续。"),
            "DRAW_AUX": ui("【辅助线模式】: 请画出透视视场下的 4 条轮廓线段，将利用交比自动计算视场中心和透视中点。")
        }
        if self.state in ('MEASURE_1', 'MEASURE_2'):
            text = self._measurement_status()
            self.status_label.setText(text)
            self.status_label.setToolTip(text)
        elif self.state in ('MARKING', 'MARK_PAUSED'):
            self.status_label.setToolTip('')
            mode = ui('已暂停') if self.state == 'MARK_PAUSED' else ui('捕捉中')
            text = (ui_format('当前帧 {p0}/{p1} 点 | {p2} | 暂停后拖动修改；Ctrl+左键可在捕捉中编辑 | 满点自动下一帧', p0=curr_len, p1=pts, p2=mode))
            if 0 <= self.selected_point < curr_len:
                d = self.marks[self.current_frame_idx][self.selected_point]
                text += ui_format(' | 已选 P{p0}（X={p1:.3f}, Y={p2:.3f}）', p0=self.selected_point+1, p1=d['rel_x'], p2=d['rel_y'])
            self.status_label.setText(ui(text))
        else:
            self.status_label.setToolTip('')
            self.status_label.setText(ui_format('状态: {p0}', p0=msgs.get(self.state, self.state)))

    def set_state(self, state):
        self.finish_point_drag()
        self.state = state
        self._refresh_state_status()
        self._sync_capture_controls()
        if state != "DRAW_AUX":
            self.aux_current_points.clear()
        if state not in ["MEASURE_1", "MEASURE_2", "IDLE"]:
            self.measure_line = None
            self.measure_points = []
            self.measure_completed = False 
        self.render_frame()

    def clear_aux_lines(self):
        self.aux_lines.clear()
        self.aux_current_points.clear()
        self.aux_intersections.clear()
        self.aux_midpoints.clear()
        self.render_frame()

    def _get_intersection(self, p1, p2, p3, p4):
        x1, y1 = p1; x2, y2 = p2
        x3, y3 = p3; x4, y4 = p4
        den = (x1 - x2)*(y3 - y4) - (y1 - y2)*(x3 - x4)
        if abs(den) < 1e-5: return None
        px = ((x1*y2 - y1*x2)*(x3 - x4) - (x1 - x2)*(x3*y4 - y3*x4)) / den
        py = ((x1*y2 - y1*x2)*(y3 - y4) - (y1 - y2)*(x3*y4 - y3*x4)) / den
        return (px, py)

    def on_video_moved(self, x, y):
        if not (x == -1.0 and y == -1.0):
            snap_dist = 15.0 / self.video_widget.zoom_level
            min_d = float('inf')
            snap_pt = None
            
            all_snap_points = self.aux_intersections + self.aux_midpoints
            for pt in all_snap_points:
                d = math.hypot(x - pt[0], y - pt[1])
                if d < snap_dist and d < min_d:
                    min_d = d
                    snap_pt = pt
                        
            if snap_pt: self.mouse_curr_pos = snap_pt
            else: self.mouse_curr_pos = (x, y)
        else:
            self.mouse_curr_pos = None
        if not self._cursor_timer.isActive():
            self._cursor_timer.start(16)

    def render_frame(self):
        if self.base_qimage is None: return

        key = (self.current_frame_idx, self._marks_revision, self.trail_mode)
        if self._overlay_cache_key != key:
            def make_marks(pts):
                return [{'pixel_pt': d['pixel_pt'], 'text': str(i+1),
                         'color': TRACK_COLORS_VIEW[i % len(TRACK_COLORS_VIEW)]}
                        for i, d in enumerate(pts)]
            current_marks = make_marks(self.marks.get(self.current_frame_idx, []))
            keys = sorted(k for k, pts in self.marks.items() if pts)
            pos = bisect_left(keys, self.current_frame_idx)
            previous_marks = make_marks(self.marks[keys[pos-1]]) if pos else []
            past = keys[:pos + (1 if pos < len(keys) and keys[pos] == self.current_frame_idx else 0)]
            show_keys = [] if self.trail_mode == 0 else (past[-10:] if self.trail_mode == 1 else keys)
            count = max((len(self.marks[k]) for k in show_keys), default=0)
            all_trails = [[] for _ in range(count)]
            for k in show_keys:
                for i, d in enumerate(self.marks[k]):
                    all_trails[i].append(d['pixel_pt'])
            self._overlay_cache = current_marks, previous_marks, all_trails
            self._overlay_cache_key = key
        current_marks, previous_marks, all_trails = self._overlay_cache
        video_info = None
        if self.video_path and not self._is_fullscreen_video:
            video_info = {'path': self.video_path, 'resolution': f'{self.video_w}×{self.video_h}',
                          'fps': f'{self.fps:.2f}', 'frames': self.total_frames,
                          'duration': f'{self.total_frames / max(self.fps, 1.0):.2f}s'}

        display_coord_origin = (
            self.coord_pending_origin
            if self.state == "SET_COORD_2" and self.coord_pending_origin is not None
            else self.coord_origin
        )

        self.video_widget.set_overlay_state(
            mode=self.state,
            calib_pt1=self.calib_pt1,
            calib_line=self.calib_line,
            coord_origin=display_coord_origin,
            coord_angle=self.coord_angle,
            measure_pt1=self.measure_pt1,
            measure_line=self.measure_line,
            measure_points=self.measure_points,
            measure_completed=self.measure_completed,
            current_marks=current_marks,
            previous_marks=previous_marks,
            mouse_curr_pos=self.mouse_curr_pos,
            video_info=video_info,
            aux_lines=self.aux_lines,
            aux_current_points=self.aux_current_points,
            aux_intersections=self.aux_intersections,
            aux_midpoints=self.aux_midpoints,
            all_trails=all_trails,
            scale_factor=self.scale_factor,
            physics_unit=self.physics_unit,
            grid_size=self.grid_size,
            show_overlays=self.show_overlays,
            cursor_mode=self.cursor_mode,
            selected_point=self.selected_point,
            marks_revision=(self._marks_revision, self.current_frame_idx, self.trail_mode)
        )
        self._sync_capture_controls()

    def get_formatted_coords(self, pt):
        is_calibrated = (self.calib_line is not None)
        if is_calibrated and self.coord_origin:
            dx = pt[0] - self.coord_origin[0]
            dy = self.coord_origin[1] - pt[1]
            theta = self.coord_angle
            rx = (dx * math.cos(theta) + dy * math.sin(theta)) * self.scale_factor
            ry = (-dx * math.sin(theta) + dy * math.cos(theta)) * self.scale_factor
            return f"{rx:.2f}, {ry:.2f} {self.physics_unit}"
        elif is_calibrated:
            return f"{pt[0]*self.scale_factor:.2f}, {pt[1]*self.scale_factor:.2f} {self.physics_unit}"
        else:
            return f"{int(pt[0])}, {int(pt[1])} px"

    def on_video_clicked(self, x, y):
        if self._seek_pending or self.current_frame_idx < 0 or self.play_timer.isActive():
            return
        self._finish_adjustment_history()
        self._last_canvas_click = None
        if self.state == 'MARKING':
            previous_entry = self.undo_stack[-1] if self.undo_stack else None
            previous_redo = list(self.redo_stack)
            self.append_capture_point(x, y)
            entry = self.undo_stack[-1] if self.undo_stack else None
            if entry is not previous_entry and entry and entry.get('mark_change'):
                self._last_canvas_click = (time.monotonic(), entry, previous_redo)
            return
        before = self._capture_history_state()
        initial_mode = self.state
        old_entry = self.undo_stack[-1] if self.undo_stack else None
        previous_redo = list(self.redo_stack)
        old_marks = copy.deepcopy(self.marks) if initial_mode == 'CALIBRATING_2' and self.marks else None
        self._handle_video_clicked(x, y)
        change = None
        if old_marks is not None and old_marks != self.marks:
            change = {'kind': 'all', 'before': old_marks, 'after': copy.deepcopy(self.marks)}
        self._commit_history(before, ui_format('画面点击: {p0}', p0=initial_mode), change)
        entry = self.undo_stack[-1] if self.undo_stack else None
        if initial_mode in ('MEASURE_1', 'MEASURE_2') and entry is not old_entry:
            self._last_canvas_click = (time.monotonic(), entry, previous_redo)

    def _handle_video_clicked(self, x, y):
        click_x, click_y = x, y
        # 仅辅助点吸附使用当前点击位置；不再使用可能滞后的上次鼠标坐标。
        for pt in self.aux_intersections + self.aux_midpoints:
            if math.hypot(x-pt[0], y-pt[1]) < 15 / self.video_widget.zoom_level:
                click_x, click_y = pt
                break

        if self.state == "DRAW_AUX":
            if len(self.aux_lines) == 4 and len(self.aux_current_points) == 0:
                self.clear_aux_lines()
            self.aux_current_points.append((click_x, click_y))
            if len(self.aux_current_points) == 2:
                A, B = self.aux_current_points
                self.aux_lines.append((A, B))
                self.aux_current_points.clear()
                
                if len(self.aux_lines) == 4:
                    M = [((l[0][0]+l[1][0])/2, (l[0][1]+l[1][1])/2) for l in self.aux_lines]
                    d = lambda p1, p2: math.hypot(p1[0]-p2[0], p1[1]-p2[1])
                    s1 = d(M[0], M[1]) + d(M[2], M[3]) 
                    s2 = d(M[0], M[2]) + d(M[1], M[3]) 
                    s3 = d(M[0], M[3]) + d(M[1], M[2]) 
                    
                    max_s = max(s1, s2, s3)
                    if max_s == s1: groupA, groupB = (0, 1), (2, 3)
                    elif max_s == s2: groupA, groupB = (0, 2), (1, 3)
                    else: groupA, groupB = (0, 3), (1, 2)
                    
                    self.aux_intersections.clear()
                    for i in groupA:
                        for j in groupB:
                            pt = self._get_intersection(*self.aux_lines[i], *self.aux_lines[j])
                            if pt: self.aux_intersections.append(pt)
                                
                    self.aux_midpoints.clear()
                    if len(self.aux_intersections) == 4:
                        corners = self.aux_intersections[:]
                        cx = sum(c[0] for c in corners) / 4.0
                        cy = sum(c[1] for c in corners) / 4.0
                        corners.sort(key=lambda c: math.atan2(c[1]-cy, c[0]-cx))
                        
                        self.aux_intersections = corners
                        C0, C1, C2, C3 = corners
                        O = self._get_intersection(C0, C2, C1, C3) 
                        
                        if O:
                            self.aux_midpoints.append(O) 
                            V1 = self._get_intersection(C0, C1, C2, C3) 
                            V2 = self._get_intersection(C1, C2, C3, C0) 
                            
                            def add_projective_midpoint(pA, pB, V):
                                if V:
                                    m = self._get_intersection(pA, pB, V, O)
                                    if m: self.aux_midpoints.append(m)
                                else:
                                    self.aux_midpoints.append(((pA[0]+pB[0])/2.0, (pA[1]+pB[1])/2.0))
                                    
                            add_projective_midpoint(C0, C1, V2)
                            add_projective_midpoint(C2, C3, V2)
                            add_projective_midpoint(C1, C2, V1)
                            add_projective_midpoint(C3, C0, V1)
            self.render_frame()

        elif self.state == "CALIBRATING_1":
            self.calib_pt1 = (click_x, click_y)
            self.set_state("CALIBRATING_2")
            
        elif self.state == "CALIBRATING_2":
            pixel_dist = math.hypot(click_x - self.calib_pt1[0], click_y - self.calib_pt1[1])
            if pixel_dist == 0: pixel_dist = 1.0
            
            real_dist, ok1 = QInputDialog.getDouble(self, ui("标定物理尺寸"), ui_format('像素跨度: {p0:.2f} px\n请输入实际物理长度:', p0=pixel_dist), 600.0, 0.001, 100000, 3)
            if ok1 and real_dist > 0:
                unit_str, ok2 = QInputDialog.getText(self, ui("设置单位名称"), ui("请输入单位简称 (如 mm, m):"), text="mm")
                if ok2 and unit_str.strip(): self.physics_unit = unit_str.strip()
                old_scale = self.scale_factor
                self.scale_factor = real_dist / pixel_dist
                self._rescale_capture_data(old_scale)
                self.calib_line = {
                    "p1": self.calib_pt1,
                    "p2": (click_x, click_y),
                    "text": f"{real_dist:.3f} {self.physics_unit}",
                }
                self.update_chart_labels()
                QMessageBox.information(self, ui("成功"), ui_format('标定映射建立！\n1 像素 = {p0:.4f} {p1}', p0=self.scale_factor, p1=self.physics_unit))
            self.calib_pt1 = None
            self.set_state("IDLE")
            
        elif self.state == "SET_COORD_1":
            self.coord_pending_origin = (click_x, click_y)
            self.set_state("SET_COORD_2")
            
        elif self.state == "SET_COORD_2":
            pending_origin = self.coord_pending_origin
            if pending_origin is None:
                self.set_state(self.coord_return_state)
                return

            dx = click_x - pending_origin[0]
            dy = pending_origin[1] - click_y 
            
            if math.hypot(dx, dy) < 3.0: 
                angle_deg = 0.0
            else: 
                angle_deg = math.degrees(math.atan2(dy, dx))
                
            final_deg, ok = QInputDialog.getDouble(
                self, ui("精调坐标轴角度"), 
                ui("你已点击定义了X轴方向。\n在此微调逆时针旋转角度(度)或直接确认\n(0 为默认水平向右):"), 
                angle_deg, -360, 360, 2
            )
            
            if ok:
                self.coord_origin = pending_origin
                self.coord_angle = math.radians(final_deg)
                # 新坐标系仅用于后续捕捉；历史记录保留捕捉时的物理坐标。
                QMessageBox.information(self, ui("设置成功"), 
                                        ui_format('自定义坐标系已生效！\n原点: ({p0:.1f}, {p1:.1f})\nX轴倾角: {p2:.2f}°', p0=self.coord_origin[0], p1=self.coord_origin[1], p2=final_deg))
            return_state = self.coord_return_state
            self.coord_pending_origin = None
            self.coord_return_state = "IDLE"
            self.set_state(return_state)
            
        elif self.state in ('MEASURE_1', 'MEASURE_2'):
            if self.state == 'MEASURE_1':
                self.measure_points = []
                self.measure_completed = False
                self.measure_line = None
            if self.measure_points and math.hypot(click_x-self.measure_points[-1][0], click_y-self.measure_points[-1][1]) < 1e-6:
                return
            self.measure_points.append((click_x, click_y))
            self.measure_pt1 = self.measure_points[0]
            self._rebuild_measurement()
            self.set_state('MEASURE_2')

    def update_chart_labels(self):
        unit = self.physics_unit if self.calib_line is not None else 'px'
        color = '#334155' if getattr(self, 'theme', 'dark') == 'light' else '#d7dde8'
        for plot, axis_name in ((self.plot_x, 'X'), (self.plot_y, 'Y')):
            plot.getAxis('left').setLabel(ui_format('{p0}轴位移', p0=axis_name), units=unit, color=color)
            plot.getAxis('bottom').setLabel(ui('相对时间'), units='s', color=color)
            for name in ('left', 'bottom'):
                axis = plot.getAxis(name)
                # pyqtgraph 0.14 的 enableAutoSIPrefix(False) 仍会计算一次前缀；
                # 禁用计算区间，保证刻度数值与单位都保持原始 px/mm/m/s。
                if hasattr(axis, 'setSIPrefixEnableRanges'):
                    axis.setSIPrefixEnableRanges(())
                axis.enableAutoSIPrefix(False)

    def update_plots(self):
        self._overlay_cache_key = None
        self._marks_revision += 1
        self._plots_dirty = True
        if not self._plot_timer.isActive():
            self._plot_timer.start(100)

    def update_highlight(self):
        pts = self.marks.get(self.current_frame_idx, [])
        selected = self.selected_point
        indices = [selected] if 0 <= selected < len(pts) else (list(range(len(pts))) if len(pts) <= 24 else [])
        xvals = [self._relative_time(self.current_frame_idx) for i in indices]
        self.highlights_x[0].setData(xvals, [pts[i]['rel_x'] for i in indices])
        self.highlights_y[0].setData(xvals, [pts[i]['rel_y'] for i in indices])

    def auto_range_plots(self):
        self.plot_x.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=True)
        self.plot_y.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=True)
        self.plot_x.autoRange()
        self.plot_y.autoRange()

    def export_current_frame(self):
        if not self.video_path or self.current_frame_idx < 0:
            QMessageBox.warning(self, ui("无法导出"), ui("请先加载视频并定位到需要导出的帧。"))
            return
        if self.play_timer.isActive(): self.toggle_play()

        default_name = f"frame_{self.current_frame_idx:06d}.png"
        filename, _ = QFileDialog.getSaveFileName(self, ui("导出当前帧原始画面"), ui(default_name), "PNG Image (*.png);;JPEG Image (*.jpg *.jpeg);;Bitmap Image (*.bmp)")
        if not filename: return
        
        lower_name = filename.lower()
        if not lower_name.endswith((".png", ".jpg", ".jpeg", ".bmp")): filename += ".png"

        # 原画导出：直接从磁盘读取 100% 原始分辨率无损导出
        cap = cv2.VideoCapture(self.video_path)
        cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame_idx)
        ret, raw_frame = cap.read()
        cap.release()

        if not ret:
            QMessageBox.critical(self, ui("导出失败"), ui("无法从视频中读取当前帧。"))
            return
        if cv2.imwrite(filename, raw_frame): QMessageBox.information(self, ui("成功"), ui_format('当前帧原始画面已导出：\n{p0}', p0=filename))
        else: QMessageBox.critical(self, ui("导出失败"), ui("写入图像文件失败，请检查文件路径或格式。"))

    def build_process_data(self, track_df):
        """由 TrackData 自动生成相邻有效帧之间的位移与速度结果。"""
        unit = self.physics_unit or "单位"
        output_columns = [
            'Time_Interval', 'Start_Frame', 'Point_Label', f'Height_Y ({unit})',
            'Delta_T (s)', f'Delta_X ({unit})', f'Velocity_X ({unit}/s)',
            f'Delta_Y ({unit})', f'Velocity_Y ({unit}/s)',
            f'True_Displacement ({unit})', f'True_Velocity ({unit}/s)'
        ]
        if track_df.empty:
            return pd.DataFrame(columns=output_columns)

        frame_col = 'Frame_Index (绝对帧数)'
        time_col = 'Absolute_Time (绝对时间-秒)'
        if frame_col not in track_df.columns or time_col not in track_df.columns:
            return pd.DataFrame(columns=output_columns)

        frame_values = pd.to_numeric(track_df[frame_col], errors='coerce').dropna()
        frame_indices = sorted(frame_values.astype(int).unique())
        if len(frame_indices) < 2:
            return pd.DataFrame(columns=output_columns)

        def point_number(column_name):
            match = re.search(r'Physical_X_P(\d+)', column_name, re.IGNORECASE)
            return int(match.group(1)) if match else 10**9

        x_cols = sorted(
            [c for c in track_df.columns if re.search(r'Physical_X_P\d+', c, re.IGNORECASE)],
            key=point_number
        )
        all_process_data = []

        for idx1, idx2 in zip(frame_indices[:-1], frame_indices[1:]):
            rows1 = track_df[pd.to_numeric(track_df[frame_col], errors='coerce') == idx1]
            rows2 = track_df[pd.to_numeric(track_df[frame_col], errors='coerce') == idx2]
            if rows1.empty or rows2.empty:
                continue
            row1, row2 = rows1.iloc[0], rows2.iloc[0]
            t1 = pd.to_numeric(pd.Series([row1[time_col]]), errors='coerce').iloc[0]
            t2 = pd.to_numeric(pd.Series([row2[time_col]]), errors='coerce').iloc[0]
            if pd.isna(t1) or pd.isna(t2):
                continue
            delta_t = float(t2 - t1)
            if delta_t <= 0:
                continue

            for x_col in x_cols:
                y_col = re.sub(r'Physical_X_', 'Physical_Y_', x_col, count=1, flags=re.IGNORECASE)
                if y_col not in track_df.columns:
                    continue
                values = pd.to_numeric(
                    pd.Series([row1[x_col], row2[x_col], row1[y_col], row2[y_col]]),
                    errors='coerce'
                )
                if values.isna().any():
                    continue
                x1, x2, y1, y2 = map(float, values)
                dx, dy = x2 - x1, y2 - y1
                vx, vy = dx / delta_t, dy / delta_t
                point_id = point_number(x_col)
                all_process_data.append({
                    'Time_Interval': f'Frame{idx1}-Frame{idx2}',
                    'Start_Frame': idx1,
                    'Point_Label': f'P{point_id}',
                    f'Height_Y ({unit})': y1,
                    'Delta_T (s)': delta_t,
                    f'Delta_X ({unit})': dx,
                    f'Velocity_X ({unit}/s)': vx,
                    f'Delta_Y ({unit})': dy,
                    f'Velocity_Y ({unit}/s)': vy,
                    f'True_Displacement ({unit})': math.hypot(dx, dy),
                    f'True_Velocity ({unit}/s)': math.hypot(vx, vy)
                })

        return pd.DataFrame(all_process_data, columns=output_columns)

    def export_to_excel(self):
        self.finish_point_drag()
        if not self.marks: return
        if self.play_timer.isActive(): self.toggle_play() 
        filename, _ = QFileDialog.getSaveFileName(self, ui("数据存档为 Excel"), ui("_tracked_data.xlsx"), "Excel Files (*.xlsx)")
        if filename:
            try:
                sorted_keys = sorted(self.marks.keys())
                data_list = []
                for k in sorted_keys:
                    pts = self.marks[k]
                    if not pts: continue
                    
                    rel_time_val = self._relative_time(k)
                    row_data = {
                        'Frame_Index (绝对帧数)': k, 
                        'Absolute_Time (绝对时间-秒)': pts[0]['time'],
                        'Relative_Time (相对时间-秒/零点平移)': rel_time_val
                    }
                    
                    for i, pt in enumerate(pts):
                        row_data[f'Physical_X_P{i+1} ({self.physics_unit})'] = pt['rel_x']
                        row_data[f'Physical_Y_P{i+1} ({self.physics_unit})'] = pt['rel_y']
                        row_data[f'Canvas_Pixel_X_P{i+1}'] = pt['pixel_pt'][0]
                        row_data[f'Canvas_Pixel_Y_P{i+1}'] = pt['pixel_pt'][1]
                        
                    data_list.append(row_data)
                df = pd.DataFrame(data_list)
                
                metadata = {
                    'Parameter': [
                        'points_per_frame', 'scale_factor', 'physics_unit', 'grid_size',
                        'coord_origin_x', 'coord_origin_y', 'coord_angle', 'first_time_sec',
                        'calib_p1_x', 'calib_p1_y', 'calib_p2_x', 'calib_p2_y', 'calib_text',
                        'video_path', 'video_w', 'video_h', 'fps', 'total_frames'
                    ],
                    'Value': [
                        self.points_per_frame, self.scale_factor, self.physics_unit, self.grid_size,
                        self.coord_origin[0] if self.coord_origin else None,
                        self.coord_origin[1] if self.coord_origin else None,
                        self.coord_angle,
                        self.first_time_sec,
                        self.calib_line['p1'][0] if self.calib_line else None,
                        self.calib_line['p1'][1] if self.calib_line else None,
                        self.calib_line['p2'][0] if self.calib_line else None,
                        self.calib_line['p2'][1] if self.calib_line else None,
                        self.calib_line['text'] if self.calib_line else None,
                        self.video_path, self.video_w, self.video_h, self.fps, self.total_frames
                    ]
                }
                df_meta = pd.DataFrame(metadata)
                process_df = self.build_process_data(df)
                
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='TrackData', index=False)
                    df_meta.to_excel(writer, sheet_name='Metadata', index=False)
                    basis_rows = []
                    for frame_index, points in self.marks.items():
                        for point_index, point in enumerate(points):
                            basis = point.get('basis')
                            if basis:
                                basis_rows.append({'Frame': frame_index, 'Point': point_index+1,
                                                   'Origin_X': basis['origin'][0], 'Origin_Y': basis['origin'][1],
                                                   'Angle': basis['angle'], 'Scale': basis['scale']})
                    if basis_rows:
                        pd.DataFrame(basis_rows).to_excel(writer, sheet_name='PointBasis', index=False)
                    process_df.drop(columns=['Start_Frame']).to_excel(
                        writer, sheet_name='ProcessData', index=False
                    )
                QMessageBox.information(self, ui("成功"), ui_format('报告成功写入到本地：\n{p0}', p0=filename))
            except Exception as e:
                QMessageBox.critical(self, ui("导出失败"), ui(str(e)))
    
    def import_from_excel(self):
        self.finish_point_drag()
        filename, _ = QFileDialog.getOpenFileName(self, ui("导入 Excel 捕捉数据"), ui(""), "Excel Files (*.xlsx)")
        if not filename: return
        if self.play_timer.isActive(): self.toggle_play()

        self._finish_adjustment_history()
        history_before = self._capture_history_state()
        marks_before = copy.deepcopy(self.marks)

        try:
            xls = pd.ExcelFile(filename)
            sheet_names = xls.sheet_names

            if 'Metadata' in sheet_names:
                df_meta = pd.read_excel(filename, sheet_name='Metadata')
                meta_dict = dict(zip(df_meta['Parameter'], df_meta['Value']))
                
                if 'points_per_frame' in meta_dict and pd.notna(meta_dict['points_per_frame']):
                    self.points_per_frame = int(meta_dict['points_per_frame'])
                    self._init_plot_curves(self.points_per_frame)
                if 'scale_factor' in meta_dict and pd.notna(meta_dict['scale_factor']):
                    self.scale_factor = float(meta_dict['scale_factor'])
                if 'physics_unit' in meta_dict and pd.notna(meta_dict['physics_unit']):
                    self.physics_unit = str(meta_dict['physics_unit'])
                if 'grid_size' in meta_dict and pd.notna(meta_dict['grid_size']):
                    self.grid_size = float(meta_dict['grid_size'])
                
                ox = meta_dict.get('coord_origin_x', meta_dict.get('origin_x'))
                oy = meta_dict.get('coord_origin_y', meta_dict.get('origin_y'))
                if pd.notna(ox) and pd.notna(oy):
                    self.coord_origin = (float(ox), float(oy))
                    
                c_angle = meta_dict.get('coord_angle')
                if pd.notna(c_angle):
                    self.coord_angle = float(c_angle)
                    
                if 'first_time_sec' in meta_dict and pd.notna(meta_dict['first_time_sec']):
                    self.first_time_sec = float(meta_dict['first_time_sec'])

                calib_keys = ('calib_p1_x', 'calib_p1_y', 'calib_p2_x', 'calib_p2_y')
                if all(k in meta_dict for k in calib_keys):
                    vals = [meta_dict[k] for k in calib_keys]
                    if all(pd.notna(v) for v in vals):
                        calib_text = meta_dict.get('calib_text', None)
                        if pd.notna(calib_text):
                            self.calib_line = {
                                'p1': (float(vals[0]), float(vals[1])),
                                'p2': (float(vals[2]), float(vals[3])),
                                'text': str(calib_text)
                            }

                imported_video = meta_dict.get('video_path', None)
                if imported_video and pd.notna(imported_video) and self.video_path:
                    imported_video = str(imported_video)
                    if imported_video != self.video_path:
                        reply = QMessageBox.question(
                            self, ui("视频不匹配"),
                            ui_format('导入数据来自:\n{p0}\n\n当前加载的是:\n{p1}\n\n仍要恢复标定参数并导入数据吗？', p0=imported_video, p1=self.video_path),
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                        )
                        if reply != QMessageBox.StandardButton.Yes:
                            self._rollback_uncommitted_action(history_before, marks_before)
                            return

            data_sheet = 'TrackData' if 'TrackData' in sheet_names else [s for s in sheet_names if s != 'Metadata'][0]
            df = pd.read_excel(filename, sheet_name=data_sheet)

            phys_x_cols = [col for col in df.columns if 'physical_x' in col.lower()]
            num_points_in_file = len(phys_x_cols) if phys_x_cols else 1
            if num_points_in_file > self.points_per_frame:
                self.points_per_frame = num_points_in_file
                self._init_plot_curves(self.points_per_frame)

            frame_col = time_col = rel_time_col = None
            for col in df.columns:
                cl = col.lower()
                if frame_col is None and 'frame' in cl: frame_col = col
                elif time_col is None and 'absolute_time' in cl: time_col = col
                elif rel_time_col is None and 'relative_time' in cl: rel_time_col = col

            if frame_col is None:
                QMessageBox.critical(self, ui("导入失败"), ui("无法识别帧索引列 (Frame_Index)"))
                self._rollback_uncommitted_action(history_before, marks_before)
                return

            self.marks = {}
            skipped = 0
            for _, row in df.iterrows():
                frame_idx = int(row[frame_col])
                if self.total_frames > 0 and not (0 <= frame_idx < self.total_frames):
                    skipped += 1
                    continue

                t = float(row[time_col]) if time_col is not None and pd.notna(row.get(time_col)) else (frame_idx / max(self.fps, 1.0))
                pts_list = []
                for i in range(self.points_per_frame):
                    px_c = next((c for c in df.columns if re.search(rf'^physical_x_p{i+1}(?:\s|$)', c, re.I) or (i==0 and c.lower() == 'physical_x (mm)')), None)
                    py_c = next((c for c in df.columns if re.search(rf'^physical_y_p{i+1}(?:\s|$)', c, re.I) or (i==0 and c.lower() == 'physical_y (mm)')), None)
                    pix_x_c = next((c for c in df.columns if re.fullmatch(rf'canvas_pixel_x_p{i+1}', c, re.I) or (i==0 and c.lower() == 'canvas_pixel_x')), None)
                    pix_y_c = next((c for c in df.columns if re.fullmatch(rf'canvas_pixel_y_p{i+1}', c, re.I) or (i==0 and c.lower() == 'canvas_pixel_y')), None)
                    
                    if px_c is None or py_c is None or pd.isna(row.get(px_c)) or pd.isna(row.get(py_c)):
                        # 旧版/新版捕捉均按顺序连续编号；缺失尾部不填零。
                        break
                    rx = float(row[px_c]) if px_c is not None and pd.notna(row.get(px_c)) else 0.0
                    ry = float(row[py_c]) if py_c is not None and pd.notna(row.get(py_c)) else 0.0
                    pix_x = float(row[pix_x_c]) if pix_x_c is not None and pd.notna(row.get(pix_x_c)) else 0.0
                    pix_y = float(row[pix_y_c]) if pix_y_c is not None and pd.notna(row.get(pix_y_c)) else 0.0
                    
                    pts_list.append({ 'time': t, 'rel_x': rx, 'rel_y': ry, 'pixel_pt': (pix_x, pix_y) })

                self.marks[frame_idx] = pts_list

            if self.first_time_sec is None and self.marks and rel_time_col is not None:
                first_key = min(self.marks.keys())
                rel_t = float(df.loc[df[frame_col] == first_key, rel_time_col].iloc[0])
                if pd.notna(rel_t):
                    self.first_time_sec = self.marks[first_key][0]['time'] - rel_t

            if 'PointBasis' in sheet_names:
                for _, basis_row in pd.read_excel(filename, sheet_name='PointBasis').iterrows():
                    frame, point = int(basis_row['Frame']), int(basis_row['Point'])-1
                    if frame in self.marks and 0 <= point < len(self.marks[frame]):
                        self.marks[frame][point]['basis'] = {
                            'origin': (float(basis_row['Origin_X']), float(basis_row['Origin_Y'])),
                            'angle': float(basis_row['Angle']), 'scale': float(basis_row['Scale'])}
            if self.first_time_sec is None and self.marks:
                self.first_time_sec = min(self.marks)/max(self.fps, 1e-9)
            self.selected_point = -1
            self.update_chart_labels()
            self.update_plots()
            self._flush_plots()
            self.auto_range_plots()
            self.render_frame()

            msg = ui_format('成功导入 {p0} 帧记录 (每帧 {p1} 点)', p0=len(self.marks), p1=self.points_per_frame)
            if skipped > 0: msg += ui_format('\n（跳过了 {p0} 条超出当前视频帧范围的记录）', p0=skipped)
            QMessageBox.information(self, ui("导入成功"), ui(msg))

            mark_change = {
                'kind': 'all',
                'before': marks_before,
                'after': copy.deepcopy(self.marks),
            }
            self._commit_history(history_before, '导入 Excel 数据', mark_change)

        except Exception as e:
            self._rollback_uncommitted_action(history_before, marks_before)
            QMessageBox.critical(self, ui("导入失败"), ui_format('读取 Excel 文件出错：\n{p0}', p0=str(e)))

    def closeEvent(self, event):
        """退出时优雅结束后台解码线程"""
        if self.play_timer.isActive():
            self.play_timer.stop()
        QApplication.instance().removeEventFilter(self)
        self.decoder_worker.stop()
        event.accept()

    def _flush_plots(self):
        if not self._plots_dirty:
            return
        if self.plot_panel.isHidden():
            return  # 隐藏图表保持脏标记，展开后补刷。
        self._plots_dirty = False
        keys = [k for k in sorted(self.marks) if self.marks[k]]
        # 按颜色合并不同点的时间序列，以 NaN 分隔；图元数与点数无关。
        groups = [([], [], [], []) for _ in TRACK_COLORS_X]
        count = max((len(self.marks[k]) for k in keys), default=0)
        for i in range(count):
            ts, xs, ys, data = groups[i % len(groups)]
            for k in keys:
                pts = self.marks[k]
                if i < len(pts):
                    d = pts[i]
                    ts.append(self._relative_time(k))
                    xs.append(d['rel_x']); ys.append(d['rel_y'])
                    data.append({'frame': k, 'point': i})
                else:
                    ts.append(float('nan')); xs.append(float('nan')); ys.append(float('nan')); data.append(None)
            ts.append(float('nan')); xs.append(float('nan')); ys.append(float('nan')); data.append(None)
        for i, (ts, xs, ys, data) in enumerate(groups):
            self.curves_x[i].setData(ts, xs, data=data, connect='finite')
            self.curves_y[i].setData(ts, ys, data=data, connect='finite')
        self.update_highlight()

    def video_cursor_update(self):
        # 鼠标移动只更新动态光标，不重算轨迹或图表。
        self.video_widget.overlay_state['mouse_curr_pos'] = self.mouse_curr_pos
        self.video_widget.update()

    def _sync_capture_controls(self):
        active = self.state in ('MARKING', 'MARK_PAUSED') and self.current_frame_idx >= 0
        self.act_pause_capture.setEnabled(active)
        paused = self.state == 'MARK_PAUSED'
        self.act_pause_capture.setText(ui('继续捕捉' if paused else '暂停捕捉'))
        self.act_pause_capture.setIcon(self._custom_icon('play' if paused else 'pause'))
        pts = self.marks.get(self.current_frame_idx, [])
        if not 0 <= self.selected_point < len(pts):
            self.selected_point = -1
        self._refresh_state_status()

    def toggle_capture_pause(self):
        if self.state not in ('MARKING', 'MARK_PAUSED'):
            return
        self.finish_point_drag()
        self._finish_adjustment_history()
        before = self._capture_history_state()
        if self.play_timer.isActive():
            self.toggle_play()
        self.set_state('MARK_PAUSED' if self.state == 'MARKING' else 'MARKING')
        self._commit_history(before, '暂停/继续捕捉')

    def select_point_by_number(self, number):
        self.select_capture_point(number - 1)
        self.finish_point_drag()

    def select_capture_point(self, index):
        self.finish_point_drag()
        if self.play_timer.isActive() or self._seek_pending:
            return
        pts = self.marks.get(self.current_frame_idx, [])
        self.selected_point = index if 0 <= index < len(pts) else -1
        if self.selected_point >= 0 and self.state in ('MARKING', 'MARK_PAUSED'):
            self._finish_adjustment_history()
            self._drag_context = {'frame': self.current_frame_idx, 'index': index,
                                  'point': copy.deepcopy(pts[index]), 'state': self._capture_history_state()}
        self.update_highlight()
        self.render_frame()

    def _make_point(self, x, y, old=None):
        if old is None:
            basis = {'origin': self.coord_origin, 'angle': self.coord_angle, 'scale': self.scale_factor}
        else:
            basis = old.get('basis')
            if basis is None:
                # 兼容旧版 Excel：以原坐标值为锚，仅转换拖动增量。
                px, py = old['pixel_pt']
                dx, dy = x-px, py-y
                c, s = math.cos(self.coord_angle), math.sin(self.coord_angle)
                result = copy.deepcopy(old)
                result.update(pixel_pt=(x, y),
                              rel_x=old['rel_x']+(dx*c+dy*s)*self.scale_factor,
                              rel_y=old['rel_y']+(-dx*s+dy*c)*self.scale_factor)
                return result
        ox, oy = basis['origin']
        dx, dy = x-ox, oy-y
        c, s = math.cos(basis['angle']), math.sin(basis['angle'])
        return {'time': old['time'] if old else self.current_frame_idx/self.fps,
                'pixel_pt': (float(x), float(y)), 'rel_x': (dx*c+dy*s)*basis['scale'],
                'rel_y': (-dx*s+dy*c)*basis['scale'], 'basis': copy.deepcopy(basis)}

    def append_capture_point(self, x, y):
        if self.coord_origin is None:
            QMessageBox.warning(self, ui('尚未设置坐标'), ui('请先设置坐标原点和 X 轴方向。'))
            return
        pts = self.marks.get(self.current_frame_idx, [])
        if len(pts) >= self.points_per_frame:
            self.set_state('MARK_PAUSED')
            return  # 已满帧不再清空重录；暂停后直接编辑。
        before = self._capture_history_state()
        exists = self.current_frame_idx in self.marks
        d = self._make_point(x, y)
        if self.first_time_sec is None:
            self.first_time_sec = d['time']
            self.update_chart_labels()
        pts = self.marks.setdefault(self.current_frame_idx, [])
        index = len(pts)
        pts.append(d)
        self.selected_point = index
        change = {'kind': 'point', 'operation': 'append', 'frame_idx': self.current_frame_idx,
                  'index': index, 'before': None, 'after': copy.deepcopy(d), 'before_exists': exists}
        self.update_plots()
        self.render_frame()
        self._refresh_state_status()
        advance = (len(pts) >= self.points_per_frame and self.auto_advance
                   and self.current_frame_idx + self.spin_step.value() < self.total_frames)
        if len(pts) >= self.points_per_frame:
            if not advance:
                self.set_state('MARK_PAUSED')
        self._commit_history(before, ui_format('新增 P{p0}', p0=index+1), change)
        if advance:
            self.seek_frame(self.spin_step.value())

    def preview_point_drag(self, x, y):
        ctx = self._drag_context
        if ctx is None or ctx['frame'] != self.current_frame_idx:
            return
        self.marks[ctx['frame']][ctx['index']] = self._make_point(x, y, ctx['point'])
        self.update_plots()
        self.render_frame()

    def finish_point_drag(self):
        ctx, self._drag_context = self._drag_context, None
        if ctx is None:
            return
        after = copy.deepcopy(self.marks[ctx['frame']][ctx['index']])
        if after != ctx['point']:
            self._commit_history(ctx['state'], ui_format('移动 P{p0}', p0=ctx['index']+1),
                                 {'kind': 'point', 'operation': 'replace', 'frame_idx': ctx['frame'],
                                  'index': ctx['index'], 'before': ctx['point'], 'after': after})
        self.update_highlight()

    def create_quality_dialog(self):
        self.quality_dialog = QDialog(self)
        self.quality_dialog.setWindowTitle(ui('画面调节'))
        self.quality_dialog.setMinimumWidth(430)
        form = QFormLayout(self.quality_dialog)
        form.setVerticalSpacing(9)
        self._quality_numbers = {}
        self._adjustment_sliders = {}
        original = (
            ('brightness', '亮度', -100, 100, 0, 1.0, '整体加亮或减暗。'),
            ('contrast', '对比度 (%)', 25, 250, 100, 100.0, '调整灰度差异。'),
            ('sharpness', '清晰度 (%)', 0, 200, 0, 100.0, '通过反锐化掩模增强边缘细节。'))
        self._adjustment_specs = original + EXTRA_ADJUSTMENT_SPECS
        for name, label, low, high, value, divisor, tip in self._adjustment_specs:
            row = QWidget(); layout = QHBoxLayout(row); layout.setContentsMargins(0, 0, 0, 0)
            slider = QSlider(Qt.Orientation.Horizontal)
            slider.setRange(low, high); slider.setValue(value); slider.setToolTip(ui(tip))
            number = QLabel(); number.setMinimumWidth(48); number.setAlignment(Qt.AlignmentFlag.AlignRight)
            slider.valueChanged.connect(lambda v, n=name: self._update_adjustment_number(n))
            slider.valueChanged.connect(self.on_adjustment_changed)
            setattr(self, 'slider_'+name, slider)
            self._adjustment_sliders[name] = slider
            self._quality_numbers[name] = number
            layout.addWidget(slider); layout.addWidget(number)
            form.addRow(ui(label), row)
        self._sync_quality_values()
        self.quality_dialog.finished.connect(lambda _: self._finish_adjustment_history())
        row = QWidget(); layout = QHBoxLayout(row); layout.setContentsMargins(0, 4, 0, 0)
        layout.addStretch()
        self.btn_reset_adjustments = QPushButton(ui('恢复默认'))
        self.btn_reset_adjustments.setFixedSize(88, 29)
        self.btn_reset_adjustments.setStyleSheet('QPushButton { min-height: 0; padding: 3px 8px; }')
        self.btn_reset_adjustments.clicked.connect(self.reset_quality_with_history)
        layout.addWidget(self.btn_reset_adjustments)
        form.addRow(row)
        note = QLabel(ui('仅调整视频显示；捕捉坐标和原始帧导出保持不变。'))
        note.setWordWrap(True)
        form.addRow(note)

    def _sync_quality_values(self):
        for name in self._quality_numbers:
            self._update_adjustment_number(name)

    def reset_quality_with_history(self):
        self._finish_adjustment_history()
        before = self._capture_history_state()
        self.reset_adjustments()
        self._commit_history(before, '恢复默认画面')

    def create_file_browser(self):
        self.browser_root = os.getcwd()
        self.file_panel = QFrame(self.centralWidget())
        self.file_panel.setObjectName('filePeekPanel')
        self.file_panel.setStyleSheet('QFrame#filePeekPanel { background: #151c25; border: 1px solid #263241; }')
        layout = QVBoxLayout(self.file_panel)
        layout.setContentsMargins(1, 1, 7, 1)
        layout.setSpacing(0)
        self.file_resize_grip = QWidget(self.file_panel)
        self.file_resize_grip.setCursor(Qt.CursorShape.SizeHorCursor)
        self.file_resize_grip.setToolTip(ui("拖动调整文件树宽度"))
        self.file_resize_grip.setMouseTracking(True)
        self.file_model = QFileSystemModel(self)
        self.file_model.setReadOnly(True)
        self.file_model.setFilter(QDir.Filter.AllDirs | QDir.Filter.Files | QDir.Filter.NoDotAndDotDot)
        patterns = ['*'+ext for ext in VIDEO_EXTENSIONS] + ['*'+ext.upper() for ext in VIDEO_EXTENSIONS]
        self.file_model.setNameFilters(patterns)
        self.file_model.setNameFilterDisables(False)
        self.file_tree = QTreeView(self.file_panel)
        self.file_tree.setStyleSheet('''
            QTreeView { background: #111821; border: none; outline: 0; }
            QTreeView::item:selected { background: #263241; color: #edf2f8; }
        ''')
        self.file_tree.verticalScrollBar().setStyleSheet('''
            QScrollBar:vertical {
                background: #090d12; width: 11px; margin: 0; border: none;
            }
            QScrollBar::handle:vertical {
                background: #bfc5cc; min-height: 28px; border-radius: 4px;
                margin: 2px;
            }
            QScrollBar::handle:vertical:hover { background: #e2e5e9; }
            QScrollBar::handle:vertical:pressed { background: #f0f2f4; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0; border: none; background: transparent;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: #090d12;
            }
        ''')
        self.file_tree.setModel(self.file_model)
        self.file_model.setRootPath(self.browser_root)
        self.file_tree.setRootIndex(QModelIndex())
        self.file_tree.setHeaderHidden(True)
        self.file_tree.setAnimated(False)
        self.file_tree.setUniformRowHeights(True)
        self.file_tree.setEditTriggers(QTreeView.EditTrigger.NoEditTriggers)
        for i in range(1, self.file_model.columnCount()):
            self.file_tree.hideColumn(i)
        self.file_tree.doubleClicked.connect(self.open_tree_video)
        self._browser_loaded_paths = set()
        self.file_model.directoryLoaded.connect(self._on_browser_directory_loaded)
        self._pending_browser_path = None
        self.set_browser_folder(self.browser_root)
        layout.addWidget(self.file_tree)
        self._position_file_browser()
        self.file_panel.hide()
        self.setMouseTracking(True)
        for widget in self.findChildren(QWidget):
            widget.setMouseTracking(True)
        QApplication.instance().installEventFilter(self)


    def set_browser_folder(self, folder):
        self.browser_root = os.path.abspath(folder)
        self._pending_browser_path = os.path.normcase(os.path.normpath(self.browser_root))
        self.file_tree.collapseAll()
        self.file_model.setRootPath(self.browser_root)
        self.file_tree.setRootIndex(QModelIndex())  # Windows: C:/、D:/、E:/…，不限制在当前目录。
        self._expand_browser_path()

    def open_tree_video(self, index):
        path = self.file_model.filePath(index)
        if os.path.isfile(path) and os.path.splitext(path)[1].lower() in VIDEO_EXTENSIONS:
            self.open_video_path(path)

    def open_video_path(self, path):
        if self.marks:
            answer = QMessageBox.question(self, ui('切换视频'), ui('切换视频将清除当前捕捉记录及撤销历史。请先导出需要保留的数据。继续打开？'),
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                          QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.finish_point_drag()
        self._finish_adjustment_history()
        if self.play_timer.isActive():
            self.toggle_play()
        self._seek_pending = True
        self.total_frames = 0
        self.current_frame_idx = -1
        self.base_qimage = None
        self.video_widget.set_frame(None)
        self.selected_point = -1
        self._drag_context = None
        self.set_browser_folder(os.path.dirname(os.path.abspath(path)))
        self.status_label.setText(ui('状态: 正在后台打开视频…'))
        self.decoder_worker.open_video(path)

    def toggle_video_fullscreen(self, checked=False):
        checked = bool(checked)
        if checked == self._is_fullscreen_video:
            return
        self.finish_point_drag()
        self._finish_adjustment_history()
        if checked:
            self._fullscreen_restore = {
                'window_state': self.windowState(),
                'widgets': [(w, not w.isHidden()) for w in
                            (self.file_panel, self.plot_panel, self.main_toolbar, self.menuBar(),
                             self.playback_bar, self.status_panel)],
                'margins': self.root_layout.contentsMargins(), 'spacing': self.root_layout.spacing(),
                'zoom': (self.video_widget.zoom_level, self.video_widget.pan_offset_x, self.video_widget.pan_offset_y)}
            for w, _ in self._fullscreen_restore['widgets']:
                w.hide()
            self.quality_dialog.hide()
            self._set_file_browser_visible(False)
            self.root_layout.setContentsMargins(0, 0, 0, 0)
            self.root_layout.setSpacing(0)
            self._is_fullscreen_video = True
            self.showFullScreen()
            QTimer.singleShot(0, self.video_widget.fit_to_view)
        else:
            restore = self._fullscreen_restore
            self._is_fullscreen_video = False
            self.setWindowState(restore['window_state'])
            for w, visible in restore['widgets']:
                w.setVisible(visible)
            self.root_layout.setContentsMargins(restore['margins'])
            self.root_layout.setSpacing(restore['spacing'])
            self.video_widget.zoom_level, self.video_widget.pan_offset_x, self.video_widget.pan_offset_y = restore['zoom']
            self._flush_plots()
        self._position_edge_handles()
        self.render_frame()

    def exit_video_fullscreen(self):
        self.toggle_video_fullscreen(False)

    def on_video_double_clicked(self, first_press_was_click=False):
        # Qt 双击之前会先发送一次单击。捕捉模式撤回这次临时加点，
        # 避免进入全屏时多出一个点；常规单击仍立即响应，不引入双击等待。
        candidate = self._last_canvas_click
        self._last_canvas_click = None
        if first_press_was_click and candidate:
            timestamp, entry, previous_redo = candidate
            if (time.monotonic()-timestamp <= QApplication.doubleClickInterval()/1000.0+0.15
                    and self.undo_stack and self.undo_stack[-1] is entry):
                self.undo_stack.pop()
                self.redo_stack[:] = previous_redo
                self._history_restoring = True
                try:
                    self._apply_mark_change(entry.get('mark_change'), use_after=False)
                    self._apply_history_state(entry['before'])
                    if entry.get('mark_change') and 'frame_idx' in entry['mark_change']:
                        self.seek_frame(entry['mark_change']['frame_idx'], absolute=True)
                finally:
                    self._history_restoring = False
                self._update_history_actions()
        if not self._is_fullscreen_video:
            self.toggle_video_fullscreen(True)

    def _position_file_browser(self):
        parent = self.centralWidget()
        limit = max(140, parent.width()-80)
        width = min(max(180, self._browser_width), limit)
        self.file_panel.setGeometry(0, 0, width, parent.height())
        self.file_resize_grip.setGeometry(width-7, 0, 7, parent.height())
        self.file_resize_grip.raise_()
        self.file_panel.raise_()
        self._position_edge_handles()



    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.Show and isinstance(watched, QMenu):
            self._translate_ui_objects([watched]+watched.findChildren(QWidget)+watched.findChildren(QAction))
        if hasattr(self, 'file_panel'):
            kind = event.type()
            if watched is self.file_resize_grip:
                if kind == QEvent.Type.MouseButtonPress and event.button() == Qt.MouseButton.LeftButton:
                    self._browser_resize_start = (event.globalPosition().x(), self.file_panel.width())
                    return True
                if kind == QEvent.Type.MouseMove and self._browser_resize_start is not None:
                    x, width = self._browser_resize_start
                    self._browser_width = max(180, min(int(width+event.globalPosition().x()-x), self.centralWidget().width()-80))
                    self._position_file_browser()
                    return True
                if kind == QEvent.Type.MouseButtonRelease and event.button() == Qt.MouseButton.LeftButton:
                    self._browser_resize_start = None
                    return True
            if kind == QEvent.Type.Resize and watched is self.centralWidget():
                self._position_file_browser()
        return super().eventFilter(watched, event)

    def _update_adjustment_number(self, name):
        value = self._adjustment_sliders[name].value()
        self._quality_numbers[name].setText(ui(f'{value/100:.2f}' if name in ('gamma', 'exposure') else str(value)))

    def _set_adjustment_controls(self):
        values = dict(self.extra_adjustments, brightness=self.brightness, contrast=self.contrast, sharpness=self.sharpness)
        for name, _, _, _, _, divisor, _ in self._adjustment_specs:
            slider = self._adjustment_sliders[name]
            slider.blockSignals(True)
            slider.setValue(int(round(values[name]*divisor)))
            slider.blockSignals(False)
        self._sync_quality_values()

    def _send_adjustments(self, update_frame=True):
        self.decoder_worker.set_adjustments(self.brightness, self.contrast, self.sharpness,
            self.current_frame_idx if update_frame else None, extra_adjustments=self.extra_adjustments)

    def _expand_browser_path(self):
        if self._pending_browser_path is None:
            return
        target = self.file_model.index(self.browser_root)
        if not target.isValid():
            return
        chain = []
        index = target
        while index.isValid():
            chain.append(index)
            index = index.parent()
        for index in reversed(chain):
            self.file_tree.expand(index)
        self.file_tree.setCurrentIndex(target)
        self.file_tree.scrollTo(target, QTreeView.ScrollHint.PositionAtCenter)
        if self._pending_browser_path in self._browser_loaded_paths:
            self._pending_browser_path = None

    def _on_browser_directory_loaded(self, path):
        normalized = os.path.normcase(os.path.normpath(path))
        self._browser_loaded_paths.add(normalized)
        if self._pending_browser_path is None:
            return
        self._expand_browser_path()
        if normalized == self._pending_browser_path:
            self._pending_browser_path = None

    def _relative_time(self, frame_index):
        fps = max(self.fps, 1e-9)
        if self.first_time_sec is not None:
            origin_frame = int(round(self.first_time_sec*fps))
        else:
            origin_frame = min((k for k, pts in self.marks.items() if pts), default=0)
        return (frame_index-origin_frame)/fps

    def _rescale_capture_data(self, previous_scale):
        # 尺寸标定统一作用于已有捕捉坐标；各点原先的原点和坐标轴方向保持不变。
        for points in self.marks.values():
            for point in points:
                basis = point.get('basis')
                scale = basis['scale'] if basis else previous_scale
                ratio = self.scale_factor/max(scale, 1e-12)
                point['rel_x'] *= ratio
                point['rel_y'] *= ratio
                if basis:
                    basis['scale'] = self.scale_factor
        if self.marks:
            self.update_plots()

    def create_edge_handles(self):
        self.browser_handle = QToolButton(ui(self.centralWidget()))
        self.browser_handle.setObjectName('edgeHandle')
        self.browser_handle.setArrowType(Qt.ArrowType.RightArrow)
        self.browser_handle.setFixedSize(16, 44)
        self.browser_handle.setToolTip(ui('视频文件树：点击展开/收起，拖动右边缘调整宽度'))
        self.browser_handle.clicked.connect(lambda: self._set_file_browser_visible(not self.file_panel.isVisible()))
        self.browser_handle.setMouseTracking(True)
        self.plot_handle = QToolButton(ui(self.centralWidget()))
        self.plot_handle.setObjectName('edgeHandle')
        self.plot_handle.setFixedSize(16, 44)
        self.plot_handle.setArrowType(Qt.ArrowType.RightArrow)
        self.plot_handle.setToolTip(ui('收起右侧绘图区'))
        self.plot_handle.clicked.connect(self.toggle_plot_panel)
        self._position_edge_handles()

    def _position_edge_handles(self):
        if not hasattr(self, 'browser_handle'):
            return
        parent = self.centralWidget()
        center_y = max(0, (self.video_widget.height()-44)//2+self.video_widget.y())
        self.browser_handle.move(self.file_panel.width() if self.file_panel.isVisible() else 0, center_y)
        self.plot_handle.move(max(0, parent.width()-16), center_y)
        self.browser_handle.setArrowType(Qt.ArrowType.LeftArrow if self.file_panel.isVisible() else Qt.ArrowType.RightArrow)
        self.plot_handle.setArrowType(Qt.ArrowType.RightArrow if self.plot_panel.isVisible() else Qt.ArrowType.LeftArrow)
        self.plot_handle.setToolTip(ui('收起右侧绘图区' if self.plot_panel.isVisible() else '展开右侧绘图区'))
        for handle in (self.browser_handle, self.plot_handle):
            handle.raise_()
            handle.show()

    def _set_file_browser_visible(self, visible):
        self.file_panel.setVisible(bool(visible))
        if visible:
            self._position_file_browser()
        self._position_edge_handles()

    def toggle_plot_panel(self):
        visible = self.plot_panel.isHidden()
        self.plot_panel.setVisible(visible)
        if not self._is_fullscreen_video:
            self._plots_collapsed = not visible
        self.plot_handle.setToolTip(ui('收起右侧绘图区' if visible else '展开右侧绘图区'))
        if visible:
            self._flush_plots()
        if self._is_fullscreen_video:
            QTimer.singleShot(0, self.video_widget.fit_to_view)
        QTimer.singleShot(0, self._position_edge_handles)

    def toggle_theme(self):
        self.apply_theme('light' if self.theme == 'dark' else 'dark')

    def apply_theme(self, theme):
        self.theme = theme
        light = theme == 'light'
        self.video_widget.canvas_background = QColor('#e9eef5' if light else '#05070a')
        self.video_widget.setStyleSheet('background-color: '+self.video_widget.canvas_background.name()+'; border: 1px solid '+('#d7dfeb' if light else '#263241')+'; border-radius: 8px;')
        self.video_widget.update()
        background, foreground, border = ('#ffffff', '#253347', '#d7dfeb') if light else ('#151c25', '#d7dde8', '#263241')
        edge = f'QToolButton#edgeHandle {{ background: {background}; color: {foreground}; border: 1px solid {border}; border-radius: 3px; padding: 0; }}'
        QApplication.instance().setStyleSheet((LIGHT_STYLE if light else APP_STYLE)+edge)
        self.file_panel.setStyleSheet(f'QFrame#filePeekPanel {{ background: {background}; border: 1px solid {border}; }}')
        tree_background = '#ffffff' if light else '#111821'
        selected = '#deebfc' if light else '#263241'
        self.file_tree.setStyleSheet(f'QTreeView {{ background: {tree_background}; color: {foreground}; border: none; outline: 0; }} QTreeView::item:selected {{ background: {selected}; color: {foreground}; }}')
        track, handle = ('#e5eaf1', '#8d99a8') if light else ('#090d12', '#bfc5cc')
        self.file_tree.verticalScrollBar().setStyleSheet(f"""
            QScrollBar:vertical {{ background: {track}; width: 11px; margin: 0; border: none; }}
            QScrollBar::handle:vertical {{ background: {handle}; min-height: 28px; border-radius: 4px; margin: 2px; }}
            QScrollBar::handle:vertical:hover {{ background: {'#6b7c90' if light else '#e2e5e9'}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; border: none; background: transparent; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: {track}; }}
        """)
        for plot, label in ((self.plot_x, 'X'), (self.plot_y, 'Y')):
            plot.setBackground('#ffffff' if light else '#101418')
            plot.setTitle(ui_format('{p0}轴位移 (依自定义坐标系)', p0=label), color=foreground)
            plot.showGrid(x=True, y=True, alpha=0.18 if light else 0.3)
            for name in ('left', 'bottom', 'right', 'top'):
                plot.getAxis(name).setPen(pg.mkPen(foreground))
                plot.getAxis(name).setTextPen(pg.mkPen(foreground))
        light_x = ['#0077b6', '#bd6800', '#477b17', '#00858b', '#8750a5', '#8e7700', '#007d95', '#20783d']
        light_y = ['#c72b59', '#cf3037', '#b45c00', '#9e005d', '#3f51b5', '#b32672', '#247b35', '#525e70']
        for curves, colors in ((self.curves_x, light_x if light else TRACK_COLORS_X), (self.curves_y, light_y if light else TRACK_COLORS_Y)):
            for curve, color in zip(curves, colors):
                curve.setPen(pg.mkPen(color, width=1.2)); curve.setSymbolBrush(color)
        for marker in self.highlights_x+self.highlights_y:
            marker.setSymbolBrush('#263447' if light else '#ffffff')
            marker.setSymbolPen(pg.mkPen('#263447' if light else '#ffffff'))
        self.act_theme.setText(ui('深色界面' if light else '浅色界面'))
        self.act_theme.setToolTip(ui('切换为深色界面' if light else '切换为浅色界面'))
        pixmap = QPixmap(20, 20); pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(foreground), 1.4)); painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawEllipse(QRectF(5, 5, 10, 10))
        for i in range(8):
            a = i*math.pi/4
            painter.drawLine(QPointF(10+7*math.cos(a), 10+7*math.sin(a)), QPointF(10+9*math.cos(a), 10+9*math.sin(a)))
        painter.end(); self.act_theme.setIcon(QIcon(pixmap))
        for action, icon in (
                (self.act_load, 'load'), (self.act_undo, 'undo'), (self.act_redo, 'redo'),
                (self.act_calib, 'calib'), (self.act_coord, 'origin'), (self.act_aux_line, 'aux'),
                (self.act_measure, 'measure'), (self.act_mark, 'mark'),
                (self.act_import, 'import'), (self.act_export, 'export'),
                (self.act_export_frame, 'frame'), (self.act_reset_view, 'reset')):
            action.setIcon(self._custom_icon(icon))
        self.btn_cursor.setIcon(self._custom_icon('cursor'))
        self.btn_trail.setIcon(self._custom_icon('trail'))
        self.btn_play.setIcon(self._custom_icon('pause' if self.play_timer.isActive() else 'play'))
        self.btn_prev.setIcon(self._custom_icon('prev'))
        self.btn_next.setIcon(self._custom_icon('next'))
        self._sync_capture_controls()
        self.update_chart_labels()

    def toggle_language(self):
        self.set_language('en' if self.language == 'zh' else 'zh')

    def set_language(self, language):
        global CURRENT_LANGUAGE
        self.language = language
        CURRENT_LANGUAGE = language
        app = QApplication.instance()
        if not hasattr(self, '_qt_translation'):
            self._qt_translation = QTranslator(self)
            self._qt_translation.load('qtbase_zh_CN', QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath))
        app.removeTranslator(self._qt_translation)
        if language == 'zh':
            app.installTranslator(self._qt_translation)
        # Existing captions and tooltips use their canonical text for either direction.
        objects = self.findChildren(QWidget)+self.findChildren(QAction)
        for obj in objects:
            if isinstance(obj, (QAction, QAbstractButton, QLabel, QMenu)):
                if hasattr(obj, 'text') and hasattr(obj, 'setText'):
                    text = obj.text()
                    canonical = UI_EN_CN.get(text, text)
                    obj.setText(ui(canonical))
                if isinstance(obj, QMenu):
                    canonical = UI_EN_CN.get(obj.title(), obj.title())
                    obj.setTitle(ui(canonical))
            for getter, setter in [('toolTip', 'setToolTip'), ('statusTip', 'setStatusTip'), ('windowTitle', 'setWindowTitle')]:
                if hasattr(obj, getter) and hasattr(obj, setter):
                    value = getattr(obj, getter)()
                    getattr(obj, setter)(ui(UI_EN_CN.get(value, value)))
        for plot in (self.plot_x,self.plot_y):
            for menu in (plot.plotItem.ctrlMenu,plot.plotItem.vb.menu):
                self._translate_ui_objects([menu]+menu.findChildren(QWidget)+menu.findChildren(QAction))
        self.spin_step.setSuffix(ui(' 帧'))
        self.act_language.setText(ui('English' if language == 'zh' else '中文'))
        self.act_language.setToolTip(ui('Switch to English' if language == 'zh' else '切换为中文'))
        for plot, axis in ((self.plot_x, 'X'), (self.plot_y, 'Y')):
            color = '#253347' if self.theme == 'light' else '#d7dde8'
            plot.setTitle(ui_format('{p0}轴位移 (依自定义坐标系)', p0=axis), color=color)
        self._sync_capture_controls()
        self.update_chart_labels()
        self.render_frame()

    def start_measurement(self):
        self.finish_point_drag()
        self._finish_adjustment_history()
        if self.play_timer.isActive():
            self.toggle_play()
        before = self._capture_history_state()
        self.measure_points = []
        self.measure_line = None
        self.measure_pt1 = None
        self.measure_completed = False
        self.set_state('MEASURE_1')
        self.video_widget.setFocus()
        self._commit_history(before, ui('开始折线测量'))

    def _rebuild_measurement(self):
        points = list(self.measure_points)
        scale = self.scale_factor if self.calib_line is not None else 1.0
        unit = self.physics_unit if self.calib_line is not None else 'px'
        lengths = [math.hypot(b[0]-a[0], b[1]-a[1])*scale for a,b in zip(points,points[1:])]
        angles = []
        for a,b,c in zip(points,points[1:],points[2:]):
            ux,uy = a[0]-b[0],a[1]-b[1]
            vx,vy = c[0]-b[0],c[1]-b[1]
            if math.hypot(ux,uy)*math.hypot(vx,vy) < 1e-12:
                angles.append(None)
            else:
                angles.append(math.degrees(math.atan2(abs(ux*vy-uy*vx),ux*vx+uy*vy)))
        self.measure_line = {'points': points, 'segments': lengths, 'angles': angles,
                             'total': sum(lengths), 'unit': unit, 'completed': self.measure_completed,
                             'p1': points[0] if points else None, 'p2': points[-1] if points else None,
                             'dist_str': f'{sum(lengths):.2f} {unit}'}

    def finish_measurement(self):
        if self.state not in ('MEASURE_1','MEASURE_2') or not self.measure_points or self.measure_completed:
            return
        self._finish_adjustment_history()
        before = self._capture_history_state()
        if len(self.measure_points) < 2:
            self.measure_points = []
            self.measure_line = None
            self.measure_pt1 = None
            self.measure_completed = False
        else:
            self.measure_completed = True
            self._rebuild_measurement()
        self.set_state('MEASURE_1')
        self._commit_history(before, ui('结束折线测量'))

    def _measurement_status(self):
        if not self.measure_points:
            return ui('折线测量：逐点点击；右键或 Enter 结束本次测量。')
        data = self.measure_line
        unit = data['unit']
        pieces = [ui_format('测量：{p0} 点 | 总长 {p1:.2f} {p2}', p0=len(self.measure_points), p1=data['total'], p2=unit)]
        if data['segments']:
            pieces.append(ui('各段')+': '+', '.join(f'L{i+1}={v:.2f} {unit}' for i,v in enumerate(data['segments'])))
        if data['angles']:
            pieces.append(ui('内夹角')+': '+', '.join(f'P{i+2}={v:.2f}°' if v is not None else f'P{i+2}=—' for i,v in enumerate(data['angles'])))
        pieces.append(ui('已结束；左键开始下一次测量。') if self.measure_completed else ui('右键或 Enter 结束。'))
        return ' | '.join(pieces)

    def _translate_ui_objects(self, objects):
        for obj in objects:
            if isinstance(obj,(QAction,QAbstractButton,QLabel)) and hasattr(obj,'text'):
                obj.setText(ui(UI_EN_CN.get(obj.text(),obj.text())))
            if isinstance(obj,QMenu):
                obj.setTitle(ui(UI_EN_CN.get(obj.title(),obj.title())))


# ==========================================================
# 4. 程序入口
# ==========================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
