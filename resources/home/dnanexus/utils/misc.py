import importlib
from pathlib import Path
from types import ModuleType
from typing import Optional

# conditional path depending on whether the script is run on DNAnexus or not
if Path("/home/dnanexus").exists():
    CONFIG_PATH = Path("/home/dnanexus/configs")
else:
    CONFIG_PATH = Path("resources/home/dnanexus/configs")


def select_config(name_config: str) -> Optional[ModuleType]:
    """Given a config name, import the appropriate module for writing the sheet

    Parameters
    ----------
    name_config : str
        Name of the config module to import

    Returns
    -------
    Optional[ModuleType]
        Module to use for config information for writing the sheet or None if
        the names do not matches
    """

    for config in CONFIG_PATH.glob("*.py"):
        module_name = config.name.replace(".py", "")

        if name_config.lower() == module_name:
            return importlib.import_module(
                f".{module_name}", f"{config.parts[-2]}"
            )

    return None
