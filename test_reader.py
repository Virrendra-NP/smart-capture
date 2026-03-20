import sys

try:
    from mpxj.reader import ProjectReader
    print("ProjectReader imported")
except Exception as e:
    import traceback
    traceback.print_exc()
