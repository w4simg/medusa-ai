import os
import sys
import time
import shutil
import subprocess
import logging
from typing import Optional
import httpx

import config

logger = logging.getLogger("MedusaServer")


class LlamaServerManager:
    """Manages the lifecycle of the local Medusa llama-server backend."""

    def __init__(self):
        self.host = config.LLAMA_SERVER_HOST
        self.port = config.LLAMA_SERVER_PORT
        self.url = config.LLAMA_SERVER_URL
        self.model_path = config.MODEL_PATH
        self.process: Optional[subprocess.Popen] = None
        self._owned = False

    def find_executable(self) -> Optional[str]:
        """Locates the llama-server executable."""
        found = shutil.which("llama-server")
        if found:
            return found

        winget_path = os.path.expandvars(
            r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\ggml.llamacpp_Microsoft.Winget.Source_8wekyb3d8bbwe\llama-server.exe"
        )
        if os.path.exists(winget_path):
            return winget_path

        return None

    def is_running(self) -> bool:
        """Checks if llama-server is currently healthy and responding."""
        try:
            with httpx.Client(timeout=2.0) as client:
                res = client.get(f"{self.url}/health")
                return res.status_code == 200
        except Exception:
            return False

    def start(self, timeout_sec: int = 40) -> bool:
        """Starts llama-server if not already running."""
        if self.is_running():
            logger.info(f"Medusa server is already running at {self.url}.")
            return True

        if not os.path.exists(self.model_path):
            logger.warning(f"Medusa model file not found at: {self.model_path}")
            return False

        exe = self.find_executable()
        if not exe:
            logger.warning("llama-server executable not found in PATH or WinGet location.")
            return False

        cmd = [
            exe,
            "-m", str(self.model_path),
            "-ngl", str(config.GPU_LAYERS),
            "-c", str(config.CONTEXT_SIZE),
            "-np", str(config.PARALLEL_SLOTS),
            "--port", str(self.port),
            "--host", str(self.host),
        ]

        logger.info(f"Starting Medusa local server: {' '.join(cmd)}")
        try:
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0,
            )
            self._owned = True

            start_time = time.time()
            while time.time() - start_time < timeout_sec:
                if self.process.poll() is not None:
                    out = self.process.stdout.read() if self.process.stdout else ""
                    logger.warning(f"Medusa server exited early: {out[:200]}")
                    self.stop()
                    return False

                if self.is_running():
                    logger.info("Medusa local server is ready and active!")
                    return True

                time.sleep(1.0)

            logger.warning("Medusa server startup timed out.")
            self.stop()
            return False
        except Exception as e:
            logger.warning(f"Failed to launch Medusa server: {e}")
            self.stop()
            return False

    def stop(self):
        """Stops the llama-server if it was started by this manager."""
        if self._owned and self.process:
            logger.info("Stopping Medusa local server...")
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except Exception:
                self.process.kill()
            finally:
                self.process = None
                self._owned = False
            logger.info("Medusa local server stopped.")
