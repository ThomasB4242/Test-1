#!/usr/bin/env python3
"""
Run the survey reporter through Claude.

Usage:
    python claude_runner.py
    python claude_runner.py "Generate a report from sample_data/Example - toplines.xlsx"
"""

import anyio
import sys
from claude_agent_sdk import query, ClaudeAgentOptions, ResultMessage, SystemMessage

SYSTEM_PROMPT = """You are a survey report assistant. You help users generate PowerPoint
reports from survey data files (.sav SPSS files or .xlsx Excel toplines).

The project is a Python survey reporting tool. Key commands:
- Generate a report: python main.py <input_file> [--output output.pptx] [--annotations file.yaml]
- Run the WA Energy example: python wa_energy_report.py
- Generate annotations template: python main.py <file> --generate-annotations annotations.yaml

Available sample files in sample_data/:
- "Example - toplines.xlsx" — formatted Excel toplines
- "sample_toplines.xlsx" — simple toplines
- "sample_survey.sav" — SPSS data file
- "wa_energy_annotations.yaml" — annotations for the WA energy report

When the user asks you to generate a report, use Bash to run the appropriate command
and report what was produced. Always use the Bash tool to actually run the commands.
"""


async def main():
    prompt = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else input("What would you like to do? ")

    print(f"\nClaude is working on: {prompt}\n")
    print("-" * 60)

    async for message in query(
        prompt=prompt,
        options=ClaudeAgentOptions(
            cwd="/home/user/Test-1",
            allowed_tools=["Bash", "Read", "Glob"],
            permission_mode="acceptEdits",
            system_prompt=SYSTEM_PROMPT,
            max_turns=10,
        ),
    ):
        if isinstance(message, ResultMessage):
            print("\n" + "-" * 60)
            print(message.result)


anyio.run(main)
