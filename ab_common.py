"""Shared constants and data model for the A/B testing pipeline.

Used by generate_data.py, validate_data.py, and analyze_ab_test.py so all
three scripts agree on what a valid campaign-recipient record looks like.
"""

from datetime import date

from pydantic import BaseModel, Field, model_validator

# The two arms of the split. "control" got the usual email; "treatment" got the
# variant being tested (a discount code in the demo data).
GROUPS = ["control", "treatment"]

# Customer segments, used only to demonstrate the "slicing until something is
# significant" pitfall - the segment breakdown is NOT the headline result.
SEGMENTS = ["new", "returning"]

RAW_DATA_PATH = "data/campaign_raw.csv"
VALIDATED_DATA_PATH = "data/campaign_validated.csv"
VALIDATION_REPORT_PATH = "data/validation_report.txt"
OUTPUT_WORKBOOK_PATH = "output/ab_test_analysis.xlsx"

CSV_FIELDS = [
    "customer_id",
    "email",
    "group",
    "segment",
    "send_date",
    "opened",
    "clicked",
    "converted",
    "order_value",
]


class CampaignRecipient(BaseModel):
    """One recipient of the campaign email and what they did afterwards.

    The funnel must be internally consistent: you cannot click an email you
    never opened, you cannot convert without clicking, and an order value only
    exists when there was an order.
    """

    customer_id: str = Field(min_length=1)
    email: str = Field(min_length=1)
    group: str = Field(pattern="^(control|treatment)$")
    segment: str = Field(pattern="^(new|returning)$")
    send_date: date
    opened: bool
    clicked: bool
    converted: bool
    order_value: float = Field(ge=0)

    @model_validator(mode="after")
    def check_funnel_consistency(self) -> "CampaignRecipient":
        if self.clicked and not self.opened:
            raise ValueError("clicked is True but opened is False")
        if self.converted and not self.clicked:
            raise ValueError("converted is True but clicked is False")
        if self.converted and self.order_value <= 0:
            raise ValueError("converted is True but order_value is not positive")
        if not self.converted and self.order_value != 0:
            raise ValueError("order_value is non-zero but converted is False")
        return self

    class Config:
        extra = "ignore"
