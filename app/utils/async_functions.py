from app.models.ASFT_Data import ASFT_Data
import concurrent.futures


def concurrent_ASFT(*pdfs: str):
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(pdfs)) as executor:
        futures = [executor.submit(ASFT_Data, pdf) for pdf in pdfs]
        results = [future.result() for future in futures]
    return results
