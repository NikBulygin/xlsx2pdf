from typing import Dict, Any

def main(params: Dict[str, Any]) -> \
    Dict[str, Any]:
    return {
        "just_variable": "some_string",
        "table": {
            '1': {'123', '123', '123', '123'},
            '2': {'123', '123', '123', '123'}
        }
    }