@echo off
cd /d "%~dp0"

echo ============================================
echo   退税申报自动化工具
echo ============================================
echo.

where node >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Node.js，请先安装
    pause
    exit /b 1
)

echo [提示] 请确保 Chrome 已打开逐项配单页面（调试端口 9222）
echo.
echo ============================================
echo  选择运行模式：
echo.
echo  [C] 继续处理（推荐）
echo     自动跳过已处理完成的行，从第一行未处理的位置继续
echo     适用于上次运行中断、需要继续的场景
echo.
echo  [N] 重新开始
echo     从第一行开始处理，清空运行状态
echo     适用于 Excel 状态已重置、想从头开始
echo.
echo  [R] 从指定行开始
echo     指定从哪一行开始处理
echo     适用于已知需要从某行继续的场景
echo ============================================
echo.

choice /C CNR /M "请输入选择 [C/N/R]"

if errorlevel 3 goto resume_row
if errorlevel 2 goto new_start
goto continue_run

:continue_run
echo.
node tax_refund_matcher.js --run
goto end

:new_start
echo.
echo [模式] 重新开始，清空状态...
del /q state\tax_refund_state.json 2>nul
node tax_refund_matcher.js --run
goto end

:resume_row
echo.
set /p row_num="请输入行号（例如 2）："
if "%row_num%"=="" set row_num=2
echo [模式] 从第 %row_num% 行开始处理...
node tax_refund_matcher.js --run --row=%row_num%
goto end

:end
echo.
pause
