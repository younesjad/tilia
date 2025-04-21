import sys
import streamlit.web.cli as stcli

sys.argv = ["streamlit", "run", "tilia_simulator.py"]
sys.exit(stcli.main())