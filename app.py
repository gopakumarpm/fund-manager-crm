import streamlit as st
import sys
import os

st.set_page_config(
    page_title="Fund Manager CRM",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("Fund Manager CRM - Debug Mode")
st.write(f"Python version: {sys.version}")
st.write(f"Streamlit version: {st.__version__}")
st.write(f"Working directory: {os.getcwd()}")
st.write(f"File location: {os.path.dirname(os.path.abspath(__file__))}")

# Test imports one by one
errors = []

try:
    import pandas as pd
    st.success(f"pandas {pd.__version__} - OK")
except Exception as e:
    errors.append(f"pandas: {e}")
    st.error(f"pandas: {e}")

try:
    import openpyxl
    st.success(f"openpyxl {openpyxl.__version__} - OK")
except Exception as e:
    errors.append(f"openpyxl: {e}")
    st.error(f"openpyxl: {e}")

try:
    from openpyxl.utils import get_column_letter
    st.success("openpyxl.utils.get_column_letter - OK")
except Exception as e:
    errors.append(f"openpyxl.utils: {e}")
    st.error(f"openpyxl.utils: {e}")

try:
    import plotly.express as px
    import plotly.graph_objects as go
    st.success(f"plotly {px.__version__} - OK")
except Exception as e:
    errors.append(f"plotly: {e}")
    st.error(f"plotly: {e}")

try:
    from datetime import datetime, date
    st.success("datetime - OK")
except Exception as e:
    errors.append(f"datetime: {e}")
    st.error(f"datetime: {e}")

try:
    import json
    import io
    import traceback
    st.success("json, io, traceback - OK")
except Exception as e:
    errors.append(f"json/io/traceback: {e}")
    st.error(f"json/io/traceback: {e}")

# Test data directory
try:
    DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    os.makedirs(DATA_DIR, exist_ok=True)
    test_file = os.path.join(DATA_DIR, "test.txt")
    with open(test_file, "w") as f:
        f.write("test")
    os.remove(test_file)
    st.success(f"Data directory writable: {DATA_DIR}")
except Exception as e:
    st.warning(f"Data dir not writable: {e}")
    import tempfile
    DATA_DIR = os.path.join(tempfile.gettempdir(), "fund_manager_data")
    os.makedirs(DATA_DIR, exist_ok=True)
    st.success(f"Using temp dir: {DATA_DIR}")

if not errors:
    st.balloons()
    st.success("All imports and tests passed! The full app should work.")
else:
    st.error(f"Found {len(errors)} errors. See above for details.")
