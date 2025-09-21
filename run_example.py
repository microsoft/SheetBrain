#!/usr/bin/env python3
# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""
Standalone example runner for SheetBrain.

This script demonstrates how to use SheetBrain from outside the package.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from core.agent import SheetBrain
# from config.settings import Config  # Uncomment if using custom configuration  


def main():
    """Example usage of the SheetBrain library."""

    print("🚀 [Example] Initializing SheetBrain...")

    # Example configuration
    excel_path = "example_table.xlsx"
    user_question = "What is the total landings (tonnes live weight) for Scotland in 2023, and how does it compare to the total landings for England, Wales, and N.I.?"

    # Check if Excel file exists
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        print("Please make sure the example Excel file is in the correct location.")
        return

    # Option 1: Use default configuration
    agent = SheetBrain(excel_path=excel_path, total_token_budget=5000)

    # Option 2: Use custom configuration (commented out)
    # config = Config(
    #     max_turns=5,
    #     enable_validation=True,
    #     enable_understanding=True
    # )
    # agent = SheetBrain(excel_path=excel_path, config=config)

    print("📋 [Example] Starting analysis...")

    try:
        # Run the analysis
        result = agent.run(
            user_question=user_question,
            max_turns=3,
            enable_validation=True,
            enable_understanding=True
        )

        print("\n" + "="*60)
        print("FINAL RESULT")
        print("="*60)
        print(f"Success: {result['success']}")
        print(f"Total Iterations: {result['total_iterations']}")
        print(f"Final Answer: {result['answer']}")
        print(f"Confidence Score: {result['confidence_score']:.2f}")
        print(f"Validation Passed: {result['validation_passed']}")
        print(f"Total Duration: {result['total_duration']:.2f}s")

        if result['issues_found']:
            print(f"\nIssues Found:")
            for issue in result['issues_found']:
                print(f"  - {issue}")

        print("="*60)

    except Exception as e:
        print(f"❌ Error running analysis: {str(e)}")
        print("Make sure you have the required dependencies installed and your API key is configured.")


if __name__ == "__main__":
    main()