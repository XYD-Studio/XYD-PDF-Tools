import PyInstaller.__main__

print("============== 开始执行商业级深度学习打包 ==============")
print("正在收集所有依赖并构建打包环境，耗时较长，请耐心等待...")

PyInstaller.__main__.run([
    'main.py',
    '-n', 'XYD_PDF_Tools_Pro',
    '-D',
    '-w',  # 隐藏控制台 (如果你想看黑框输出排错，可以先用 # 注释掉这行)
    '-i', 'logo.ico',

    # 静态资源
    '--add-data', 'logo.ico;.',
    '--add-data', 'gs_portable;gs_portable',

    # 把 public 文件夹及其里面的图片打包进去！
    '--add-data', 'public;public',

    # 强制收集依赖
    '--collect-all', 'Cython',
    '--collect-all', 'paddle',
    '--collect-all', 'paddleocr',
    '--collect-all', 'shapely',
    '--collect-all', 'pyclipper',
    '--collect-all', 'skimage',
    '--collect-all', 'scipy',
    '--collect-all', 'cv2',
    '--collect-all', 'lmdb',
    '--collect-all', 'imgaug',  # 显式收集 imgaug
    '--collect-all', 'imageio',  # 显式收集 imageio

    # 【最关键的救命参数】：强行保留第三方库的版本元数据，防止运行时 importlib 崩溃！
    '--copy-metadata', 'imageio',
    '--copy-metadata', 'imgaug',
    '--copy-metadata', 'paddleocr',
    '--copy-metadata', 'scikit-image',
    '--copy-metadata', 'numpy',
    '--copy-metadata', 'Pillow',

    # 隐式引入容易被静态分析忽略的内部依赖
    '--hidden-import', 'paddleocr',
    '--hidden-import', 'tools',
    '--hidden-import', 'ppocr',
    '--hidden-import', 'imghdr',
    '--hidden-import', 'imgaug',
    '--hidden-import', 'imageio',

    # 忽略不必要的巨型库，加快打包速度，减小体积
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'tkinter',
])

print("============== 打包完成 ==============")