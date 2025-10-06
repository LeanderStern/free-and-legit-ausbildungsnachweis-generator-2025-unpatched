from datetime import datetime
from functools import cached_property

from base_model import PydanticBaseModel


class WorkDay(PydanticBaseModel):
    date: datetime
    hours_worked: float
    text: str
    location: str | None

    @cached_property
    def normalized_day_name(self) -> str:
        return self.date.strftime("%A").lower()