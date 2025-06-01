# prl_launcher.py
import subprocess
import time
import webbrowser
from pathlib import Path

def launch_streamlit_app():
    app_script = Path(__file__).parent / "appy1.py"

    # Start Streamlit via subprocess, as if typing `streamlit run appy1.py`
    subprocess.Popen(
        ["streamlit", "run", str(app_script), "--server.headless=false"],
        shell=True
    )

    # Give Streamlit a few seconds to start
    time.sleep(3)

    # Open the browser to the local Streamlit app
    webbrowser.open("http://localhost:8501")

    input("✅ Streamlit app launched! Press ENTER to exit this launcher...")

if __name__ == "__main__":
    launch_streamlit_app()
