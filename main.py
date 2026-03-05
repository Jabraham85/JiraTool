"""Avalanche Jira Template Creator — entry point."""
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    try:
        from app import AvalancheApp
        app = AvalancheApp()
        app.run()
    except Exception:
        import traceback as _tb
        print("Fatal error while launching the app. See jira_debug.log for details.")
        _tb.print_exc()
        from utils import debug_log
        debug_log("Fatal exception launching app:\n" + traceback.format_exc())


if __name__ == "__main__":
    main()
