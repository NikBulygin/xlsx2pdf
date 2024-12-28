from logger import Logger
from config import Config
from templator import Templator

if __name__ == "__main__":
    Logger.setup_logging()
    Logger.print("Start report generator")
    result = Templator.start_gen(Config.arg_parser())
    Logger.print(f"Finish work report generator with {result}")
    print(result)