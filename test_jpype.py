import mpxj
import jpype
import os

try:
    print("JPype version:", jpype.__version__)
    print("MPXJ module loaded successfully")
except Exception as e:
    print(f"Error: {e}")
