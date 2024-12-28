import argparse
import json
import os
from typing import Any, Dict, Optional
from logger import Logger

class Config:
    @staticmethod
    def load_config(
        config_path:str
    ) -> Dict[str, Any]:
        if not os.path.isfile(config_path):
            Logger.print(f"Config file not found at {config_path}", level="critical")
            raise FileNotFoundError(f"")
        
        with open(config_path, "r", encoding="utf-8") as config_file:
            return json.load(config_file)
        
    @staticmethod
    def validate_path(full_path:str) -> None:
        if not os.path.isfile(full_path):
            Logger.print(f"Required file {full_path} not found", level="critical")
            raise FileNotFoundError(f"Required file {full_path} not found")
        if not os.access(full_path, os.R_OK):
            Logger.print(f"File {full_path} is not readable", level="critical")
            raise PermissionError(f"File {full_path} is not readable")
        
    @staticmethod
    def validate_out_path(out_path:str):
        out_dir = os.path.dirname(out_path)
        if not os.path.isdir(out_dir):
            Logger.print(f"Output directory {out_dir} does not exists", level="critical")
            raise FileNotFoundError(f"Output directory {out_dir} does not exists")
        if not os.access(out_dir, os.W_OK):
            Logger.print(f"Output directory {out_dir} is not writable", level="critical")
            raise PermissionError(f"Output directory {out_dir} is not writable")
        
    @staticmethod
    def validate_params(params: Dict[str, Any]) -> None:
        try:
            json.dumps(params)
        except (TypeError, ValueError) as e:
            Logger.print(f"Invalid params: {e}", level='critical')
            raise ValueError(f"Invalid params: {e}")
        
    @staticmethod
    def prepare_config(
        report: str,
        report_path: Optional[str] = None,
        path_to_py_module: Optional[str] = None,
        path_to_watermark: Optional[str] = None,
        out: Optional[str] = None,
        data_was_prepared: Optional[bool] = False,
        params: Optional[Dict[str, any]] = None,
        metadata: Optional[Dict[str, any]] = None
    ):
        config:Dict[str, Any] = Config.load_config()
        if not report_path:
            report_path = next(
                (item["path_to_report"] for item in config.get("reports", []) \
                 if item["report_name"] == report),
                 None
            )
            if not report_path:
                Logger.print(f"Report path {report} not found in config.json", level='critical')
                raise ValueError(f"Report path {report} not found in config.json")
        Config.validate_path(report_path)

        if not data_was_prepared:
            if not path_to_py_module:
                path_to_py_module = next(
                    (item["path_to_py_module"] for item in config.get("reports", []) \
                     if item["report_name"] == report),
                     None
                )
                if not path_to_py_module:
                    Logger.print(f"Path to py modyle {report} not found in config.json", level='critical')
                    raise ValueError(f"Path to py modyle {report} not found in config.json")
                Config.validate_path(path_to_py_module)

        if not path_to_watermark:
            path_to_watermark = next(
                (item["path_to_watermark"] for item in config.get("reports", []) \
                 if item["report_name"] == report),
                 None
            )
        if(path_to_watermark):
            Config.validate_path(path_to_watermark)

        if not out:
            out = config.get("default_out")
            if not out:
                Logger.print("default_out print is not found in config.json", level='critical')
                raise ValueError("default_out print is not found in config.json")
        Config.validate_out_path(out)

        libreoffice = config.get("libreoffice_calc_path")
        if not libreoffice:
            Logger.print("libreoffice_calc_path is not found in config.json", level='critical')
            raise ValueError("libreoffice_calc_path is not found in config.json")
        Config.validate_path(libreoffice)
        
        Config.validate_params(params or {})

        default_metadata: Dict[str, str] = \
            config.get("default_metadata")
        if not default_metadata:
            Logger.print("default_metadata is not found in config.json", level='warning')
        if not metadata:
            metadata = default_metadata
        else:
            metadata.update(default_metadata)
        Config.validate_params(metadata)

        return {
            "report_name": report,
            "report_path": report_path,
            "path_to_py_module": path_to_py_module,
            "path_to_watermark": path_to_watermark,
            "output": out,
            "libreoffice_path": libreoffice,
            "params": params or {},
            "metadata": metadata or {}
        }
    
    @staticmethod
    def arg_parser():
        parser = argparse.ArgumentParser(
            description=(
                "A tool to process and generate",
                "pdf reports based on xlsx templates"
            )
        )
        
        # Function for collapse code in IDE
        def add_args():
            parser.add_argument(
                "--report",
                type=str,
                required=True,
                help="The name of the report"
            )
            parser.add_argument(
                "--report_path",
                type=str,
                required=False,
                help="The path to xlsx template"
            )
            parser.add_argument(
                "--path_to_py_module",
                type=str,
                required=False,
                help="The path to extend py module"
            )
            parser.add_argument(
                "--path_to_watermark",
                type=str,
                required=False,
                help="The file path to watermark.pdf"
            )
            parser.add_argument(
                "--out",
                type=str,
                required=False,
                help="The path where output will be saved"
            )
            parser.add_argument(
                "--data_was_prepared",
                type=str,
                default="False",
                help="If true data from params put into xlsx without py module"
            )
            parser.add_argument(
                "--params",
                type=json.loads,
                required=False,
                default={},
                help={"Value for py module or directly xlsx(if data was prepared true)"}
            ),
            parser.add_argument(
                "--metadata",
                type=json.loads,
                default={},
                help="Metadata for result pdf"
            )

        add_args()
        args = parser.parse_args()
        return Config.prepare_config(
            report=args.report,
            report_path=args.report_path,
            path_to_py_module=args.path_to_py_module,
            path_to_watermath=args.path_to_watermark,
            out=args.out,
            data_was_prepared= True if str(args.data_was_prepared).lower() == "true" else False,
            params=args.params,
            metadata=args.metadata
        )
            
                    

