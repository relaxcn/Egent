@echo off
echo.
echo ===========================================
echo    Excel Agent 开发环境启动脚本
echo ===========================================
echo.

echo [1/2] 安装依赖...
call npm install

echo.
echo [2/2] 启动前端开发服务器...
start "Frontend Dev Server" cmd /k "npm run dev-server"

echo.
echo ✅ 开发环境启动完成！
echo.
echo 📍 前端服务器: https://localhost:3000
echo.
echo 现在可以运行 'npm start' 来启动Excel插件调试
echo.
pause