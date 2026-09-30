import os
import sys
import time
import webbrowser
import threading

def open_browser():
    """等待 Streamlit 服务启动后自动打开默认浏览器"""
    time.sleep(3)
    webbrowser.open("http://127.0.0.1:8505")

if __name__ == "__main__":
    # 强制禁用 developmentMode，防止 PyInstaller 打包时被误判为 Streamlit 源码开发模式
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = "8505"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    # 后台线程延迟打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()

    # 多重候选路径寻找 app.py
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    meipass = getattr(sys, "_MEIPASS", "")

    candidates = [
        os.path.join(meipass, "app.py") if meipass else "",
        os.path.join(exe_dir, "app.py"),
        os.path.join(exe_dir, "_internal", "app.py"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py"),
        os.path.join(os.getcwd(), "app.py"),
    ]

    app_path = None
    for p in candidates:
        if p and os.path.exists(p):
            app_path = os.path.abspath(p)
            break

    if not app_path:
        app_path = os.path.abspath(os.path.join(exe_dir, "app.py"))

    print("=" * 60)
    print("           FCT3 数据分析看板正在启动...")
    print("=" * 60)
    print(f"主程序脚本: {app_path}")
    print("本地访问地址: http://127.0.0.1:8505")
    print("提示：浏览器已自动打开，关闭此控制台窗口即可停止服务。")
    print("=" * 60)

    # 导入并运行 Streamlit CLI
    try:
        import streamlit.web.cli as stcli
    except ImportError:
        import streamlit.cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--server.port=8505",
        "--server.headless=true",
        "--server.address=127.0.0.1",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())
