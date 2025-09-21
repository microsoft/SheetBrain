# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Configuration settings for SheetBrain."""

import os
from typing import Optional
from dataclasses import dataclass


@dataclass
class Config:
    """Configuration class for SheetBrain application."""

    # OpenAI Configuration
    api_key: str = "your_api_key"
    base_url: str = "your_base_url"
    deployment: str = "your_model_name"

    # Processing Configuration
    max_turns: int = 3
    total_token_budget: int = 5000

    # Features
    enable_validation: bool = True
    enable_understanding: bool = True

    # Timeouts and Retries
    max_retries: int = 3
    timeout: int = 30

    @classmethod
    def from_env(cls) -> "Config":
        """Create configuration from environment variables."""
        return cls(
            api_key=os.getenv("OPENAI_API_KEY", cls.api_key),
            base_url=os.getenv("OPENAI_BASE_URL", cls.base_url),
            deployment=os.getenv("OPENAI_DEPLOYMENT", cls.deployment),
            max_turns=int(os.getenv("MAX_TURNS", cls.max_turns)),
            total_token_budget=int(os.getenv("TOKEN_BUDGET", cls.total_token_budget)),
            enable_validation=os.getenv("ENABLE_VALIDATION", "true").lower() == "true",
            enable_understanding=os.getenv("ENABLE_UNDERSTANDING", "true").lower() == "true",
            max_retries=int(os.getenv("MAX_RETRIES", cls.max_retries)),
            timeout=int(os.getenv("TIMEOUT", cls.timeout))
        )