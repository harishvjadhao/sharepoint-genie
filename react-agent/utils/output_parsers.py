from typing import Dict

from langchain_core.output_parsers import PydanticOutputParser
from pydantic import BaseModel, Field


class Summary(BaseModel):
    title: str = Field(description="title of the document")
    summary: str = Field(description="summary")

    def to_dict(self) -> Dict[str, str]:
        return {"title": self.title, "summary": self.summary}


summary_parser = PydanticOutputParser(pydantic_object=Summary)
