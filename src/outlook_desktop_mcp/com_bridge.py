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
import time
import logging

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

# How long start() waits for Outlook before letting the server come up anyway,
# and how often the COM thread retries afterwards.
_STARTUP_GRACE = 15.0
_RETRY_INTERVAL = 3.0

_NOT_CONNECTED = (
    "Outlook Desktop (Classic) is not reachable. The server is retrying in the "
    "background and will start working as soon as classic OUTLOOK.EXE is running "
    "in this session — restarting the MCP server is not required."
)


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
        self._connected = False
        self._warned_unreachable = False

    def start(self):
        """Start the COM thread. Call once at server startup.

        Never raises for an absent Outlook. A session and Outlook race each other
        at boot on any machine with autologon, and a server that exits here stays
        dead for the whole session, because MCP servers are not reloaded
        mid-session. Come up regardless and keep trying in the background.
        """
        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=_STARTUP_GRACE):
            logger.warning(
                "Outlook not reachable after %.0fs (%s). Starting anyway and "
                "retrying every %.0fs.",
                _STARTUP_GRACE, self._init_error, _RETRY_INTERVAL,
            )

    def _connect(self):
        """(Re)acquire the Outlook COM objects. Must run on the COM thread."""
        import win32com.client

        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

    def _try_connect(self) -> bool:
        """One connection attempt. COM thread only. Never raises."""
        try:
            self._connect()
            # Touch the namespace to confirm the new connection is live.
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
        except Exception as e:
            self._connected = False
            self._init_error = e
            if not self._warned_unreachable:
                # Only the first failure is loud; after that it is a poll.
                logger.warning("Outlook not reachable: %s", e)
                self._warned_unreachable = True
            else:
                logger.debug("Outlook still not reachable: %s", e)
            return False

        self._connected = True
        self._init_error = None
        self._warned_unreachable = False
        logger.info("Connected to Outlook. Store: %s, User: %s", store_name, user_name)
        self._ready.set()
        return True

    def _reconnect(self) -> bool:
        """Re-establish a dead Outlook connection. COM thread only."""
        logger.warning("Outlook COM connection lost; attempting to reconnect...")
        self._connected = False
        return self._try_connect()

    def _com_thread_main(self):
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client  # noqa: F401  (ensures COM modules load on this thread)

        pythoncom.CoInitialize()
        try:
            next_attempt = 0.0
            while not self._shutdown.is_set():
                # Keep trying to reach Outlook while idle, so a session that
                # started before Outlook did heals itself without a restart.
                if not self._connected and time.monotonic() >= next_attempt:
                    if not self._try_connect():
                        next_attempt = time.monotonic() + _RETRY_INTERVAL

                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    if not self._connected and not self._try_connect():
                        raise RuntimeError(_NOT_CONNECTED)
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
            # Only reached if the loop itself breaks, not for an absent Outlook.
            self._init_error = e
            self._ready.set()
            logger.error("COM thread failed: %s", e)
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
