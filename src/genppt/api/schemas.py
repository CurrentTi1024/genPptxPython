from pydantic import BaseModel, ConfigDict, StrictStr


class GenerationRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    template: StrictStr
    data: StrictStr
