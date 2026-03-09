try:
    import pandas as pd
except ModuleNotFoundError:
    print("Module 'pandas' is not installed. Install with: pip install pandas")
    pd = None

import docx
import win32com.client

print("OK")
