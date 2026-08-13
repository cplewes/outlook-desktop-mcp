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

# A freshly started Outlook answers DefaultStore long before its folder views are
# populated. Wait for the inbox count to hold still before serving anything.
_READY_TIMEOUT = 20.0
_READY_STABLE_FOR = 1.5
_READY_POLL = 0.5

_OL_FOLDER_INBOX = 6

_NOT_RUNNING = (
    "Classic Outlook (OUTLOOK.EXE) is not running. This server attaches to the "
    "Outlook you already have open and deliberately will not start one: an Outlook "
    "launched by automation comes up cold and answers queries from folders that are "
    "still filling, so a search can return an incomplete list with no error — which "
    "reads as 'that message does not exist'. Start Outlook Desktop (Classic) and the "
    "server will connect by itself within a few seconds; no restart needed."
)

_NOT_READY = (
    "Outlook is running but its folders have not finished loading, so any answer now "
    "could be incomplete. The server is waiting for it to settle and will serve calls "
    "as soon as it has; no restart needed."
)


class OutlookNotRunning(RuntimeError):
    """Raised when no Outlook is available to attach to."""


class OutlookNotReady(RuntimeError):
    """Raised when Outlook is attached but its stores are still populating."""


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

    def _outlook_is_running(self) -> bool:
        """True if an OUTLOOK.EXE exists in this process's Windows session.

        GetActiveObject would be the obvious check, but an Outlook that Windows
        COM-activated itself (started with -Embedding) does not publish itself to
        the running object table, so GetActiveObject reports "not running" for an
        Outlook that is plainly there. Ask the process list instead.
        """
        import os
        import win32com.client
        import win32process

        try:
            session = win32process.ProcessIdToSessionId(os.getpid())
        except Exception:
            session = None

        try:
            wmi = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")
            found = wmi.ExecQuery(
                "SELECT ProcessId, SessionId FROM Win32_Process WHERE Name='OUTLOOK.EXE'"
            )
        except Exception as e:
            # Fail open: not being able to look is not evidence Outlook is gone,
            # and refusing every call because WMI hiccuped would take the server
            # down for a check that is only about not spawning a stray Outlook.
            # The readiness gate still guards correctness either way.
            logger.warning("Could not check for a running Outlook (%s); attaching anyway.", e)
            return True

        for proc in found:
            if session is None or proc.SessionId == session:
                return True
        return False

    def _connect(self):
        """Attach to a running Outlook. Must run on the COM thread.

        Dispatch() attaches to a running Outlook, but silently COM-launches a
        fresh hidden one when there is none — and a just-launched Outlook answers
        queries from folders that are still filling, returning short result sets
        with no error. Refuse rather than start one; the README already requires
        Outlook to be running.

        There is a narrow race here: Outlook could exit between the check and the
        Dispatch, in which case we would launch one after all. The readiness gate
        covers that case rather than leaving it to serve a cold store.
        """
        import win32com.client

        if not self._outlook_is_running():
            raise OutlookNotRunning(_NOT_RUNNING)

        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

    def _wait_until_ready(self) -> bool:
        """True once the default store stops changing under us. COM thread only.

        DefaultStore.DisplayName answers immediately on an Outlook that is still
        loading, so it cannot be used to decide readiness. Item counts moving is
        the signal that folders are still being populated.
        """
        deadline = time.monotonic() + _READY_TIMEOUT
        last = None
        stable_since = None

        while time.monotonic() < deadline:
            try:
                count = self._namespace.GetDefaultFolder(_OL_FOLDER_INBOX).Items.Count
            except Exception as e:
                logger.debug("store not ready: %s", e)
                time.sleep(_READY_POLL)
                continue

            now = time.monotonic()
            if count != last:
                last = count
                stable_since = now
            elif (now - stable_since) >= _READY_STABLE_FOR:
                logger.debug("store settled at %s items", count)
                return True
            time.sleep(_READY_POLL)

        logger.warning(
            "Outlook did not settle within %.0fs (inbox count still moving, last %s).",
            _READY_TIMEOUT, last,
        )
        return False

    def _try_connect(self) -> bool:
        """One connection attempt. COM thread only. Never raises."""
        try:
            self._connect()
            # Touch the namespace to confirm the new connection is live.
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            # Liveness is not readiness. Do not serve calls until it settles.
            if not self._wait_until_ready():
                raise OutlookNotReady(_NOT_READY)
        except Exception as e:
            self._connected = False
            self._init_error = e
            if not self._warned_unreachable:
                # Only the first failure is loud; after that it is a poll.
                logger.warning("Outlook unavailable: %s", e)
                self._warned_unreachable = True
            else:
                logger.debug("Outlook still unavailable: %s", e)
            return False

        self._connected = True
        self._init_error = None
        self._warned_unreachable = False
        logger.info("Connected to Outlook. Store: %s, User: %s", store_name, user_name)
        self._ready.set()
        return True

    def _unavailable_reason(self) -> str:
        """The most specific thing we can say about why calls cannot be served."""
        if isinstance(self._init_error, (OutlookNotRunning, OutlookNotReady)):
            return str(self._init_error)
        if self._init_error:
            return f"{_NOT_RUNNING} (last error: {self._init_error})"
        return _NOT_RUNNING

    def _reconnect(self) -> bool:
        """Re-establish a dead Outlook connection. COM thread only.

        Outlook closing and reopening is the case this exists for. The reconnect
        waits for the replacement to settle before the caller's retry runs, so a
        call that spans a restart cannot come back with a half-loaded folder.
        """
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
                        raise RuntimeError(self._unavailable_reason())
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
