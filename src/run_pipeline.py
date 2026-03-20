"""
Pipeline orchestrator for PPTX deconstruct/reconstruct.
Runs all steps in sequence:
  1. deconstruct  — extract .pptx into component library
  2. generate_config — create config JSON from manifest
  3. update_config — resolve tokens from data/ folder
  4. inject — inject resolved values into _raw/ slide XML
  5. reconstruct — repack _raw/ into output .pptx

Usage:
  python run_pipeline.py <input.pptx>
  python run_pipeline.py <input.pptx> --data-dir data/ --output output.pptx --force
  python run_pipeline.py <input.pptx> --skip-deconstruct --skip-config-gen
"""

import argparse
import logging
import os
import shutil
import sys

# Ensure src/ is on the import path when run as a script
_src_dir = os.path.dirname(os.path.abspath(__file__))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

logger = logging.getLogger(__name__)


def run_step(step_name, func, **kwargs):
    """Run a pipeline step with status output."""
    logger.info("\n%s", '=' * 60)
    logger.info("  STEP: %s", step_name)
    logger.info("%s", '=' * 60)
    try:
        result = func(**kwargs)
        return result
    except SystemExit as e:
        if e.code and e.code != 0:
            logger.error("\n[FAILED] %s exited with code %s", step_name, e.code)
            sys.exit(e.code)
    except Exception as e:
        logger.error("\n[FAILED] %s: %s", step_name, e)
        sys.exit(1)


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    parser = argparse.ArgumentParser(
        description="Run the full PPTX deconstruct/reconstruct pipeline."
    )
    parser.add_argument("input_pptx", help="Path to the input .pptx file")
    parser.add_argument("--library", default="component_library",
                        help="Component library directory (default: component_library)")
    parser.add_argument("--config-dir", default="configs",
                        help="Configs output directory (default: configs)")
    parser.add_argument("--data-dir", default="data",
                        help="Data directory for token resolution (default: data)")
    parser.add_argument("--output", default="output.pptx",
                        help="Output .pptx path (default: output.pptx)")
    parser.add_argument("--force", action="store_true",
                        help="Overwrite existing library/config without prompting")
    parser.add_argument("--skip-deconstruct", action="store_true",
                        help="Skip deconstruction step (reuse existing library)")
    parser.add_argument("--skip-config-gen", action="store_true",
                        help="Skip config generation step (reuse existing config)")
    parser.add_argument("--skip-update", action="store_true",
                        help="Skip config update step (no data resolution)")
    parser.add_argument("--skip-inject", action="store_true",
                        help="Skip injection step")
    parser.add_argument("--hints", default=None,
                        help="Path to dynamic hints JSON file for config generation")
    parser.add_argument("--clean-backups", action="store_true",
                        help="Remove old backup directories from the library path")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview injection changes without modifying files")
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("--verbose", "-v", action="store_true",
                           help="Enable debug-level logging for detailed output")
    verbosity.add_argument("--quiet", "-q", action="store_true",
                           help="Suppress info messages; show only warnings and errors")
    args = parser.parse_args()

    # Apply verbosity after initial basicConfig
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    elif args.quiet:
        logging.getLogger().setLevel(logging.WARNING)

    if not os.path.exists(args.input_pptx):
        logger.error("[ERROR] Input file not found: %s", args.input_pptx)
        sys.exit(1)

    if args.clean_backups:
        import glob as _glob
        patterns = [
            f"{args.library}_backup_*",
        ]
        removed = 0
        for pattern in patterns:
            for backup_dir in _glob.glob(pattern):
                if os.path.isdir(backup_dir):
                    shutil.rmtree(backup_dir)
                    logger.info("Removed backup: %s", backup_dir)
                    removed += 1
        # Also clean _raw_clean if requested
        raw_clean = os.path.join(args.library, "_raw_clean")
        if os.path.exists(raw_clean):
            shutil.rmtree(raw_clean)
            logger.info("Removed clean backup: %s", raw_clean)
            removed += 1
        logger.info("Cleaned %d backup(s)", removed)

    # Derive config path from deck name
    deck_name = os.path.splitext(os.path.basename(args.input_pptx))[0].replace(" ", "_").lower()
    config_path = os.path.join(args.config_dir, f"{deck_name}.json")

    # ── Step 1: Deconstruct ──────────────────────────────────────────
    if not args.skip_deconstruct:
        from deconstruct import deconstruct
        run_step("Deconstruct",
                 deconstruct,
                 pptx_path=args.input_pptx,
                 library_path=args.library,
                 force=args.force)
    else:
        logger.info("\n  [skip] Deconstruct")

    # ── Step 2: Generate Config ──────────────────────────────────────
    if not args.skip_config_gen:
        from generate_config import generate_config
        run_step("Generate Config",
                 generate_config,
                 library_path=args.library,
                 configs_dir=args.config_dir,
                 force=args.force,
                 hints_file=args.hints)
    else:
        logger.info("\n  [skip] Generate Config")

    # ── Step 3: Update Config ────────────────────────────────────────
    if not args.skip_update:
        if os.path.exists(args.data_dir):
            from update_config import update_config
            run_step("Update Config",
                     update_config,
                     config_path=config_path,
                     data_dir=args.data_dir)
        else:
            logger.info("\n  [skip] Update Config -- data directory '%s' not found", args.data_dir)
    else:
        logger.info("\n  [skip] Update Config")

    # ── Step 4: Inject ───────────────────────────────────────────────
    if not args.skip_inject:
        if os.path.exists(config_path):
            from inject import inject
            run_step("Inject",
                     inject,
                     config_path=config_path,
                     library_path=args.library,
                     dry_run=args.dry_run)
        else:
            logger.info("\n  [skip] Inject -- config not found at '%s'", config_path)
    else:
        logger.info("\n  [skip] Inject")

    # ── Step 5: Reconstruct ──────────────────────────────────────────
    from reconstruct import reconstruct
    run_step("Reconstruct",
             reconstruct,
             library_path=args.library,
             output_path=args.output)

    # ── Verify ───────────────────────────────────────────────────────
    if os.path.exists(args.input_pptx) and os.path.exists(args.output):
        from reconstruct import verify
        logger.info("\n%s", '=' * 60)
        logger.info("  VERIFY")
        logger.info("%s", '=' * 60)
        verify(args.input_pptx, args.output, args.library)

    logger.info("\n%s", '=' * 60)
    logger.info("  PIPELINE COMPLETE")
    logger.info("  Output: %s", args.output)
    logger.info("%s", '=' * 60)


if __name__ == "__main__":
    main()
