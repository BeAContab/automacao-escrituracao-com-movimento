from collections.abc import Callable
from typing import Any
from loguru import logger
from time import sleep


def retry_on_exception(
    func: Callable,
    max_attempts: int = 1,
    polling_seconds: float | None = 0.5,
    exception: type[Exception] | tuple[type[Exception], ...] = Exception,
    *args,
    **kwargs,
) -> Any:
    """
    Invoke ``func``, retrying up to ``max_attempts`` times if it raises ``exception``.
    Returns whatever ``func`` returns on the first successful attempt.

    Args:
        func: Callable with zero or more arguments.
        max_attempts: Maximum number of execution attempts. Defaults to 1
            (invokes ``func`` at least once).
        exception:
            Re-invokes ``func`` if ``exception`` is raised.
        polling_seconds:
            Amount of seconds to wait before invoking ``func`` again.
            If ``polling_seconds`` is None, retries ``func`` immediately.
        *args, **kwargs:
            Any positional or keyword arguments to be passed to ``func``.
    Returns:
        Whatever ``func()`` returns on the first successful attempt.
    Raises:
        Exception: Re-raises the last caught exception (of the type(s) passed via ``exception``)
        when all attempts are exhausted.
        ValueError: If ``max_attempts`` is less than 1.
    Example:
        >>> def flaky():
        ...     ...
        >>> retry_on_exception(func=flaky, max_attempts=3)
    """
    if max_attempts < 1:
        raise ValueError(f"max_attempts must be >= 1, got {max_attempts}")

    for attempt in range(max_attempts):
        try:
            return func(*args, **kwargs)
        except exception as exc:
            f = func.__name__ if hasattr(func, "__name__") else "func"
            e = type(exc).__name__
            msg = f"Function {f} failed to run because it raised a(n) {e} error."
            if attempt == max_attempts - 1:
                logger.error(f"{msg} (Attempt {attempt + 1}/{max_attempts})")
                raise
            logger.debug(
                f"{msg} Retrying{f' in {polling_seconds} second(s)' if polling_seconds else ''}... (Attempt {attempt + 1}/{max_attempts})"
            )
            if polling_seconds:
                sleep(polling_seconds)

    return None  # If max_attempts <= 0
