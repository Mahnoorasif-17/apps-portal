import gc
import os
import time
from openpyxl import load_workbook

from .step1 import process_step_1
from .step2 import process_step_2
from .step3 import process_step_3
from .step4 import process_step_4
from .step5 import process_step_5
from .step6 import process_step_6
from .utils import *
from .netsuite_lookup import NetSuiteLookup


def _flush(wb, path):
    """Save, close, reopen — releases openpyxl internal memory."""
    wb.save(path)
    wb.close()
    del wb
    gc.collect()
    return load_workbook(path)


def run_processing_pipeline(filepath, return_output_path=False,
                             item_master_path=None, customer_path=None):
    tmp_path = filepath + ".tmp.xlsx"
    wb = None
    try:
        start_total = time.time()
        print("--- Pipeline Start ---")

        # Build NetSuite lookups once (used by Steps 4, 5, 6)
        ns_lookup = NetSuiteLookup(
            item_master_path=item_master_path,
            customer_path=customer_path,
        )

        t = time.time()
        wb = process_step_1(filepath)
        print(f"Step 1 took: {time.time() - t:.2f}s")

        t = time.time()
        process_step_2(wb)
        wb = _flush(wb, tmp_path)
        print(f"Step 2 took: {time.time() - t:.2f}s")

        t = time.time()
        process_step_3(wb)
        wb = _flush(wb, tmp_path)
        print(f"Step 3 took: {time.time() - t:.2f}s")

        t = time.time()
        process_step_4(wb, ns_lookup=ns_lookup)
        wb = _flush(wb, tmp_path)
        print(f"Step 4 took: {time.time() - t:.2f}s")

        t = time.time()
        process_step_5(wb)
        wb = _flush(wb, tmp_path)
        print(f"Step 5 took: {time.time() - t:.2f}s")

        t = time.time()
        process_step_6(wb)
        print(f"Step 6 took: {time.time() - t:.2f}s")

        new_filename = generate_new_filename(filepath)
        save_path = os.path.abspath(new_filename)
        print(f"SAVING FILE TO: {save_path}")
        wb.save(save_path)
        wb.close()
        del wb
        gc.collect()

        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except:
                pass

        print(f"--- TOTAL: {time.time() - start_total:.2f}s ---")

        if return_output_path:
            return save_path, None

    except ValidationError as ve:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except:
                pass
        return None, str(ve)
    except Exception as e:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except:
                pass
        return None, str(e)