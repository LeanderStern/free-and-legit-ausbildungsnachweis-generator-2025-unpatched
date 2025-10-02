import logging
from functools import cached_property
from logging import Logger
from typing import ClassVar

from pydantic import BaseModel, ConfigDict

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

class PydanticBaseModel(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True,
                              strict=True)
    _LOGGER: ClassVar[Logger | None] = None

    @cached_property
    def logger(self) -> Logger:
        if self._LOGGER is None:
            self.__class__._LOGGER = logging.getLogger(f"{self.__class__.__module__}.{self.__class__.__name__}")
        return self.__class__._LOGGER