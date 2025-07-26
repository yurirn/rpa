_callback = None

def set_logger_callback(callback):
    global _callback
    _callback = callback

def log_message(message, level="INFO"):
    if _callback:
        _callback(message, level)
    else:
        print(f"[{level}] {message}")