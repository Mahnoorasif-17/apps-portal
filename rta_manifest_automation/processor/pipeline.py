from .step1 import process_step_1
from .step2 import process_step_2
from .step3 import process_step_3
from .step4 import process_step_4
from .step5 import process_step_5
from .step6 import process_step_6
from .utils import *


def run_processing_pipeline(filepath, return_output_path=False):
    try:
        wb = process_step_1(filepath)
        process_step_2(wb)
        process_step_3(wb)
        process_step_4(wb)
        process_step_5(wb)
        process_step_6(wb)
        new_filename = generate_new_filename(filepath)
        wb.save(new_filename)
        if return_output_path:
            return new_filename, None
    except ValidationError as ve:
        temp_output = generate_new_filename(filepath)
        # if wb:
        #     wb.save(temp_output)
        return temp_output, str(ve)