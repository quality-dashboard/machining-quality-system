@echo off
rem ============================================================
rem 机加工质量检测系统 启动脚本
rem 端口 8505 (与GP12的8504错开, 可同时开两个系统)
rem ============================================================
cd /d D:\machining_quality_system
streamlit run app.py --server.port 8505
pause
