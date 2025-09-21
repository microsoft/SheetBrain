# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Main entry point for SheetBrain CLI."""

import argparse
import sys
from typing import Optional

from core.agent import SheetBrain
from config.settings import Config


def main():
    """Main entry point for SheetBrain CLI."""
    parser = argparse.ArgumentParser(description="SheetBrain - AI-powered Excel analysis")

    parser.add_argument("excel_path", help="Path to the Excel file to analyze")
    parser.add_argument("question", help="Question to ask about the Excel file")
    parser.add_argument("--max-turns", type=int, default=3, help="Maximum number of execution turns (default: 3)")
    parser.add_argument("--no-validation", action="store_true", help="Disable validation stage")
    parser.add_argument("--no-understanding", action="store_true", help="Disable understanding stage")
    parser.add_argument("--token-budget", type=int, default=10000, help="Token budget for context generation (default: 10000)")
    parser.add_argument("--api-key", help="OpenAI API key (overrides config)")
    parser.add_argument("--base-url", help="OpenAI base URL (overrides config)")
    parser.add_argument("--deployment", help="Model deployment name (overrides config)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")

    args = parser.parse_args()

    try:
        # Create configuration
        config = Config.from_env()

        # Override with command line arguments if provided
        if args.api_key:
            config.api_key = args.api_key
        if args.base_url:
            config.base_url = args.base_url
        if args.deployment:
            config.deployment = args.deployment

        # Initialize SheetBrain
        agent = SheetBrain(
            excel_path=args.excel_path,
            config=config,
            total_token_budget=args.token_budget
        )

        # Run analysis
        result = agent.run(
            user_question=args.question,
            max_turns=args.max_turns,
            enable_validation=not args.no_validation,
            enable_understanding=not args.no_understanding
        )

        # Print results
        print("\n" + "="*60)
        print("ANALYSIS RESULTS")
        print("="*60)
        print(f"Success: {'✅' if result['success'] else '❌'}")
        print(f"Answer: {result['answer']}")
        print(f"Confidence: {result['confidence_score']:.2f}/1.0")
        print(f"Iterations: {result['total_iterations']}")
        print(f"Duration: {result['total_duration']:.2f}s")

        if result['issues_found']:
            print(f"\nIssues Found:")
            for issue in result['issues_found']:
                print(f"  - {issue}")

        if args.verbose and result['improvement_feedback']:
            print(f"\nImprovement Feedback:")
            print(result['improvement_feedback'])

        print("="*60)

        # Exit with appropriate code
        sys.exit(0 if result['success'] else 1)

    except Exception as e:
        print(f"❌ Error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()