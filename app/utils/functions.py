from app.models.ASFT_Data import ASFT_Data
import concurrent.futures
from typing import List


def concurrent_ASFT(*pdfs: str) -> List[ASFT_Data]:
    """
    Processes multiple PDF files concurrently to extract ASFT_Data objects.

    Args:
        *pdfs (str): Variable length argument list of PDF file paths to be processed.

    Returns:
        List[ASFT_Data]: A list of ASFT_Data objects extracted from the provided PDF files.
    """
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(pdfs)) as executor:
        futures = [executor.submit(ASFT_Data, pdf) for pdf in pdfs]
        results = [future.result() for future in futures]
    return results
