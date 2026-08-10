"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.
"""
import threading
import queue
import asyncio
import sys
import logging

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")


# HRESULTs that indicate the cached Outlook COM object is dead (e.g. Outlook
# was closed and reopened) and the connection should be re-established.
_DISCONNECT_CODES = {
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
    0x800706BE,  # RPC_S_CALL_FAILED
    0x800706BF,  # RPC_S_CALL_FAILED_DNE
    0x80010108,  # RPC_E_DISCONNECTED
    0x800401FD,  # CO_E_OBJNOTCONNECTED
    0x80080005,  # CO_E_SERVER_EXEC_FAILURE
    0x80010105,  # RPC_E_SERVERFAULT
}


def _is_disconnect(e: Exception) -> bool:
    """True if the exception is a COM error signalling a dead connection."""
    hr = getattr(e, "hresult", None)
    if hr is None:
        args = getattr(e, "args", None)
        if args:
            hr = args[0]
    if not isinstance(hr, int):
        return False
    return (hr & 0xFFFFFFFF) in _DISCONNECT_CODES


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None

    def start(self):
        """Start the COM thread. Call once at server startup."""
        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=15):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                "Outlook COM thread failed to initialize within 15s. "
                "Is Outlook Desktop (Classic) running?"
            )

    def _connect(self):
        """(Re)acquire the Outlook COM objects. Must run on the COM thread."""
        import win32com.client

        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

    def _reconnect(self) -> bool:
        """Attempt to re-establish a dead Outlook connection. COM thread only."""
        try:
            logger.warning("Outlook COM connection lost; attempting to reconnect...")
            self._connect()
            # Touch the namespace to confirm the new connection is live.
            _ = self._namespace.DefaultStore.DisplayName
            logger.info("Reconnected to Outlook COM.")
            return True
        except Exception as e:
            logger.error("Reconnect to Outlook failed: %s", e)
            return False

    def _com_thread_main(self):
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client  # noqa: F401  (ensures COM modules load on this thread)

        pythoncom.CoInitialize()
        try:
            self._connect()
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = func(
                        self._outlook, self._namespace, *args, **kwargs
                    )
                except Exception as e:
                    # If Outlook was closed/reopened, the cached COM object is
                    # dead. Re-establish the connection and retry the call once.
                    if _is_disconnect(e) and self._reconnect():
                        try:
                            result_holder["value"] = func(
                                self._outlook, self._namespace, *args, **kwargs
                            )
                        except Exception as e2:
                            result_holder["error"] = e2
                    else:
                        result_holder["error"] = e
                finally:
                    result_event.set()
        except Exception as e:
            self._init_error = e
            self._ready.set()  # Unblock the caller so they see the error
            logger.error("COM thread init failed: %s", e)
        finally:
            pythoncom.CoUninitialize()

    async def call(self, func, *args, **kwargs):
        """
        Schedule a function to run on the COM thread and await its result.

        The function signature must be: func(outlook, namespace, *args, **kwargs)
        """
        result_event = threading.Event()
        result_holder = {}
        self._request_queue.put((func, args, kwargs, result_event, result_holder))

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(
            None, lambda: result_event.wait(timeout=60)
        )
        if not signaled:
            raise TimeoutError(
                "Outlook COM operation timed out after 60 seconds. "
                "Outlook may be waiting for user input (e.g., a dialog box)."
            )

        if "error" in result_holder:
            raise result_holder["error"]
        return result_holder.get("value")

    def stop(self):
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
