@echo off
chcp 65001 >nul
echo ========================================
echo 试卷分析PPT生成器 - Windows打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

echo [1/4] 安装依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [2/4] 安装PyInstaller...
pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo [3/4] 清理旧的构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo [4/4] 开始打包（单文件模式）...
pyinstaller --clean main.spec

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 可执行文件位置: dist\试卷分析PPT生成器.exe
echo.
echo 提示：首次运行可能需要几秒钟加载时间
echo.

pause