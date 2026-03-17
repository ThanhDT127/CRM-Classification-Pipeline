"""
CRM Classification Pipeline - Main Runner
==========================================
Run the full pipeline: regex → prepare LLM → call LLM → merge.

Usage:
  python run_pipeline.py           # run all steps
  python run_pipeline.py 1         # run only step 1
  python run_pipeline.py 1 2       # run steps 1 and 2
  python run_pipeline.py 3 4       # run steps 3 and 4 (LLM + merge)
"""

import sys

def main():
    steps_to_run = set()
    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            try:
                steps_to_run.add(int(arg))
            except ValueError:
                pass
    if not steps_to_run:
        steps_to_run = {1, 2, 3, 4}

    print("=" * 60)
    print("CRM Classification Pipeline")
    print(f"Steps to run: {sorted(steps_to_run)}")
    print("=" * 60)

    if 1 in steps_to_run:
        print("\n" + "─" * 40)
        print("STEP 1: Regex Classification")
        print("─" * 40)
        from step1_classify import main as step1
        step1()

    if 2 in steps_to_run:
        print("\n" + "─" * 40)
        print("STEP 2: Prepare LLM Input")
        print("─" * 40)
        from step2_prepare_llm import main as step2
        step2()

    if 3 in steps_to_run:
        print("\n" + "─" * 40)
        print("STEP 3: Call LLM (Gemini)")
        print("─" * 40)
        from step3_call_llm import main as step3
        step3()

    if 4 in steps_to_run:
        print("\n" + "─" * 40)
        print("STEP 4: Merge LLM Results")
        print("─" * 40)
        from step4_merge import main as step4
        step4()

    print("\n" + "=" * 60)
    print("✓ Pipeline finished!")
    print("=" * 60)


if __name__ == "__main__":
    main()
