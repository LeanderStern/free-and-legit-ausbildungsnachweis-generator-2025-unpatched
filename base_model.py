from pydantic import BaseModel, ConfigDict

class PydanticBaseModel(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True,
                              strict=True)