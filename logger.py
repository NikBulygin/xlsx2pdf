import datetime 
import getpass
import inspect
import logging
import os
import socket

class Logger:
    @staticmethod
    def setup_logging() -> None:
        log_dir = os.path.join(
            os.getcwd(),
            "logs"
        )
        os.makedirs(
            log_dir,
            exist_ok=True
        )
        log_file = os.path.join(
            log_dir,
            f"{datetime,datetime.now().strftime("%Y-%m-%d-%H-%M-%S")}.log"
        )

        username = getpass.getuser()
        hostname = socket.gethostname()
        ip_address = "UnknowIP"

        try:
            ip_address = socket.gethostbyname(hostname)
        except:
            ip_address = ip_address
        
        logging.basicConfig(
            filename=log_file,
            filemode='a',
            format=(
                f"%(asctime)s - %(levelname)s - User:{username},"
                f"Host:{hostname}, IP:{ip_address}\n"
                f"\t %(message)s"
            )
        )

    @staticmethod
    def print(*args, level="info", **kwargs) -> None:
        caller = inspect.stack()[1]
        caller_name = caller.function
        message = " ".join(str(arg) for arg in args)
        level = level.lower()
        if level == "debug":
            logging.debug(f"[{caller_name}] {message}")
        elif level == "info":
            logging.info(f"[{caller_name}] {message}")
        elif level == "warning":
            logging.warning(f"[{caller_name}] {message}")
        elif level == "error":
            logging.error(f"[{caller_name}] {message}")
        elif level == "critical":
            logging.critical(f"[{caller_name}] {message}")
        else:
            logging.criticla(f"Unknow logging level {level} from {caller_name}")
            logging.critical(f"[{caller_name}] {message}")
        if kwargs.pop("original_print", False):
            print(f"{message}")