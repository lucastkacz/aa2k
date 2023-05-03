from app.models.ASFT_Data import ASFT_Data
import concurrent.futures
from typing import List, Union


def concurrent_ASFT(*pdfs: Union[str, List[str]]) -> List[ASFT_Data]:
    """
    Processes multiple PDF files concurrently to extract ASFT_Data objects.

    Args:
        *pdfs (Union[str, List[str]]): Variable length argument list of PDF file paths or a list of paths to be processed.

    Returns:
        List[ASFT_Data]: A list of ASFT_Data objects extracted from the provided PDF files.
    """
    flat_pdfs = []
    for pdf in pdfs:
        if isinstance(pdf, list):
            flat_pdfs.extend(pdf)
        else:
            flat_pdfs.append(pdf)

    with concurrent.futures.ThreadPoolExecutor(max_workers=len(flat_pdfs)) as executor:
        futures = [executor.submit(ASFT_Data, pdf) for pdf in flat_pdfs]
        results = [future.result() for future in futures]
    return results
